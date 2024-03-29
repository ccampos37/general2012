VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frPDTSunatSCTR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PDT Retenciones y Contribuciones sobre Remuneraciones"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frPDTSunatSCTR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
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
      TabIndex        =   29
      Top             =   5460
      Width           =   1395
   End
   Begin VB.CommandButton cmExportar 
      Caption         =   "&Exportar"
      Height          =   390
      Left            =   5070
      TabIndex        =   28
      Top             =   4950
      Width           =   1395
   End
   Begin VB.CommandButton cmImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   5070
      TabIndex        =   27
      Top             =   4425
      Width           =   1395
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Trab."
      Height          =   390
      Left            =   5070
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label Total1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Left            =   2070
      TabIndex        =   24
      Top             =   6000
      Width           =   1305
   End
   Begin VB.Label C1 
      BackColor       =   &H80000010&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total General a Pagar"
      Height          =   285
      Index           =   7
      Left            =   105
      TabIndex        =   23
      Top             =   6000
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   6
      Left            =   3390
      TabIndex        =   22
      Top             =   5700
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   6
      Left            =   2070
      TabIndex        =   21
      Top             =   5700
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Retencion Renta 5ta. Cat."
      Height          =   285
      Index           =   6
      Left            =   105
      TabIndex        =   20
      Top             =   5700
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   5
      Left            =   3390
      TabIndex        =   19
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   5
      Left            =   2070
      TabIndex        =   18
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " EsSalud Vida"
      Height          =   285
      Index           =   5
      Left            =   105
      TabIndex        =   17
      Top             =   5400
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   4
      Left            =   3390
      TabIndex        =   16
      Top             =   5100
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   4
      Left            =   2070
      TabIndex        =   15
      Top             =   5100
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F.D.S. Artistas"
      Height          =   285
      Index           =   4
      Left            =   105
      TabIndex        =   14
      Top             =   5100
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   3
      Left            =   3390
      TabIndex        =   13
      Top             =   4800
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   3
      Left            =   2070
      TabIndex        =   12
      Top             =   4800
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Salud"
      Height          =   285
      Index           =   3
      Left            =   105
      TabIndex        =   11
      Top             =   4800
      Width           =   1950
   End
   Begin VB.Label Tri1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   2
      Left            =   3390
      TabIndex        =   10
      Top             =   4500
      Width           =   1305
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   285
      Index           =   2
      Left            =   2070
      TabIndex        =   9
      Top             =   4500
      Width           =   1305
   End
   Begin VB.Label C1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Pensiones"
      Height          =   285
      Index           =   2
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
      Caption         =   "Resultados de la Exportaci�n de Planillas para el PDT  - Formulario 0600 para la versi�n PDT Sunat 2.0"
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   615
      TabIndex        =   0
      Top             =   60
      Width           =   5550
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   15
      Picture         =   "frPDTSunatSCTR.frx":030A
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
Attribute VB_Name = "frPDTSunatSCTR"
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
            xFile = VPTAREA & "\0600" & Right(frPlans.LPlans.SelectedItem.SubItems(1), 4) & Mid(frPlans.LPlans.SelectedItem.SubItems(1), 4, 2) & REGSISTEMA.RUC & ".DJT"
            If Dir$(xFile) <> "" Then
                If MsgBox("YA EXISTE EN ESTA RUTA UN ARCHIVO CORRESPONDIENTE AL PDT SUNAT - REMUNERACIONES, DESEA UD. REEMPLAZAR EL ARCHIVO POR EL NUEVO QUE EST� PROCESANDO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
                Kill xFile
            End If
            Open xFile For Append As #1
            Do While Not RSPDT.EOF
                CADPDT = ""
                CADPDT = Val(RSPDT!TIPDOC) & "|" & RSPDT!DOCIDEN & "|" & Round(RSPDT!DIASTRAB, 0) & "|" & IIf(RSPDT!REMUIES = 0, "", RSPDT!REMUIES) & "|" & IIf(RSPDT!REMUPENSION = 0, "", RSPDT!REMUPENSION) & "|" & IIf(RSPDT!REMUSALUD = 0, "", RSPDT!REMUSALUD) & "|" & IIf(RSPDT!REMUARTISTAS = 0, "", RSPDT!REMUARTISTAS) & "|" & IIf(RSPDT!REMU5TA = 0, "", RSPDT!REMU5TA) & "|" & IIf(RSPDT!TRIBUTO5TA = 0, "", RSPDT!TRIBUTO5TA) & "|"
                Print #1, CADPDT
                RSPDT.MoveNext
            Loop
            Close #1
            RSPDT.MoveFirst
            MsgBox "PROCESO COMPLETADO. INGRESE AL PDT SUNAT Y ESCOJA LA OPCI�N IMPORTAR DEL MEN� DECLARACIONES, DENTRO DEL M�DULO 0600 DDJJ RETENCIONES Y CONTRIBUCIONES - REMUNERACIONES", vbInformation
            Exit Sub
Err1:
            MsgBox ERR.Description
            Exit Sub
End Sub

Private Sub CMIMPRIMIR_CLICK()
    Screen.MousePointer = 11
    With Reporte
        .Reset
        .WindowTitle = "PLAN0027 - REPORTE DE EXPORTACI�N DE DATOS AL PDT SUNAT"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0027.RPT"
        .Connect = "DSN=" & VGL_SERVER & ";User=sa;PWD=;DSQ=" & REGSISTEMA.BASESQL & ""
        .StoredProcParam(0) = " [##PDTSUNAT" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowState = crptMaximized
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='RUC N� " & REGSISTEMA.RUC & "'"
        .Formulas(2) = "XMES='CORRESPONDIENTE AL MES DE " & frPlans.LPlans.SelectedItem.Text & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMQUITAR_CLICK()
    If RSPDT.RecordCount = 0 Or RSPDT.EOF Then
        MsgBox "DEBE EXISTIR UN TRABAJADOR", vbCritical
        Exit Sub
    End If
    If MsgBox("REALMENTE DESEA QUITAR DE LA EXPORTACI�N PARA EL PDT SUNAT AL TRABAJADOR " & RSPDT!NOMBRES, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM PDTSUNAT WHERE CODTRAB='" & RSPDT!CODTRAB & "'"
    REFRESCARDG
End Sub

Private Sub DGLISTA_AFTERUPDATE()
    Dim RSAUX As ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT SUM(REMUIES) AS SUMA1,SUM(REMUPENSION) AS SUMA2,SUM(REMUSALUD) AS SUMA3,SUM(REMUARTISTAS) AS SUMA4,SUM(REMU5TA) AS SUMA5,SUM(TRIBUTO5TA) AS SUMA6,SUM(REMUIES*0.05) AS PAGO1,SUM(REMUPENSION*0.13) AS PAGO2,SUM(REMUSALUD*0.09) AS PAGO3,SUM(REMUARTISTAS*0.1667) AS PAGO4,SUM(ESVIDA) AS PAGO5 FROM  [##PDTSUNAT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount > 0 Then
        Tot1(1).Caption = Format(Round(RSAUX!Suma1, 0), "0.00 ")
        Tot1(2).Caption = Format(Round(RSAUX!Suma2, 0), "0.00 ")
        Tot1(3).Caption = Format(Round(RSAUX!Suma3, 0), "0.00 ")
        Tot1(4).Caption = Format(Round(RSAUX!SUMA4, 0), "0.00 ")
        Tot1(6).Caption = Format(RSAUX!SUMA5, "0.00 ")
        Tot1(5).Caption = Format(RSAUX!PAGO5 / 2, "0.00 ")
        Tri1(6).Caption = Format(Round(RSAUX!SUMA6, 0), "0.00 ")
        Tri1(5).Caption = Format(RSAUX!PAGO5, "0.00 ")
        Tri1(1).Caption = Format(Round(RSAUX!PAGO1, 0), "0.00 ")
        Tri1(2).Caption = Format(Round(RSAUX!PAGO2, 0), "0.00 ")
        Tri1(3).Caption = Format(Round(RSAUX!PAGO3, 0), "0.00 ")
        Tri1(4).Caption = Format(Round(RSAUX!PAGO4, 0), "0.00 ")
        Total1.Caption = Format(Val(Tot1(1)) + Val(Tot1(2)) + Val(Tot1(3)) + Val(Tot1(4)) + Val(Tot1(5)) + Val(Tot1(6)), "0.00 ")
        Total2.Caption = Format(Val(Tri1(1)) + Val(Tri1(2)) + Val(Tri1(3)) + Val(Tri1(4)) + Val(Tri1(5)) + Val(Tri1(6)), "0.00 ")
    End If
    Set RSAUX = Nothing
End Sub

Private Sub DGLISTA_HEADCLICK(ByVal COLINDEX As Integer)
    RSPDT.Sort = dgLista.Columns(COLINDEX).Caption
End Sub

Private Sub Form_Load()
    Set RSPDT = New ADODB.Recordset
    RSPDT.Open " [##PDTSUNAT" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic, adLockOptimistic
    REFRESCARDG
End Sub

Public Sub REFRESCARDG()
    RSPDT.Requery
    Set dgLista.DataSource = RSPDT
    With dgLista
        .Columns("NOMBRES").Width = 2600
        .Columns("NOMBRES").Locked = True
        .Columns("DIASTRAB").NumberFormat = "0.00 "
        .Columns("DIASTRAB").Alignment = dbgRight
        .Columns("REMUIES").NumberFormat = "0.00 "
        .Columns("REMUIES").Alignment = dbgRight
        .Columns("REMUPENSION").NumberFormat = "0.00 "
        .Columns("REMUPENSION").Alignment = dbgRight
        .Columns("REMUSALUD").NumberFormat = "0.00 "
        .Columns("REMUSALUD").Alignment = dbgRight
        .Columns("REMUARTISTAS").NumberFormat = "0.00 "
        .Columns("REMUARTISTAS").Alignment = dbgRight
        .Columns("REMU5TA").NumberFormat = "0.00 "
        .Columns("REMU5TA").Alignment = dbgRight
        .Columns("TRIBUTO5TA").NumberFormat = "0.00 "
        .Columns("TRIBUTO5TA").Alignment = dbgRight
        .Columns("ESVIDA").NumberFormat = "0.00 "
        .Columns("ESVIDA").Alignment = dbgRight
    End With
    DGLISTA_AFTERUPDATE
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSPDT = Nothing
End Sub

