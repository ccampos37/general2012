VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form InputPl 
   Caption         =   "X"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   Icon            =   "InputPl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   8865
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame frameContenedorx2 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   45
      TabIndex        =   25
      Top             =   6825
      Width           =   5925
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Gen.Ingresos"
         Height          =   285
         Left            =   0
         TabIndex        =   32
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label xTGIng 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   1530
         TabIndex        =   31
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label9 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Gen.Egresos"
         Height          =   285
         Left            =   2940
         TabIndex        =   30
         Top             =   0
         Width           =   1635
      End
      Begin VB.Label xTGEgr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   4590
         TabIndex        =   29
         Top             =   0
         Width           =   1395
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Gen.Neto"
         Height          =   285
         Left            =   2940
         TabIndex        =   28
         Top             =   300
         Width           =   1635
      End
      Begin VB.Label xTGNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   4590
         TabIndex        =   27
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label xNumTrabs 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 0 Trabajadores"
         Height          =   285
         Left            =   0
         TabIndex        =   26
         Top             =   300
         Width           =   2925
      End
   End
   Begin VB.Frame frameContenedorx1 
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   6105
      TabIndex        =   2
      Top             =   2490
      Width           =   2790
      Begin VB.CommandButton Command1 
         Caption         =   "&Procesar"
         Height          =   420
         Left            =   1290
         TabIndex        =   24
         Top             =   3915
         Width           =   1290
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   420
         Left            =   1290
         TabIndex        =   23
         Top             =   4395
         Width           =   1290
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   135
         Picture         =   "InputPl.frx":030A
         Top             =   750
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   120
         Picture         =   "InputPl.frx":064C
         Top             =   450
         Width           =   240
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Quitar Trab. "
         Height          =   285
         Left            =   45
         TabIndex        =   22
         Top             =   720
         Width           =   1500
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Imprimir "
         Height          =   285
         Left            =   45
         TabIndex        =   21
         Top             =   420
         Width           =   1500
      End
      Begin VB.Label xTotalCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000010&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   20
         Top             =   120
         Width           =   2640
      End
      Begin VB.Label xCodTrab 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Básico"
         Height          =   285
         Left            =   45
         TabIndex        =   19
         Top             =   1020
         Width           =   1500
      End
      Begin VB.Label xNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   1545
         TabIndex        =   18
         Top             =   3135
         Width           =   1155
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Neto Previo"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   45
         TabIndex        =   17
         Top             =   3135
         Width           =   1500
      End
      Begin VB.Label xTotEgr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   1545
         TabIndex        =   16
         Top             =   2520
         Width           =   1155
      End
      Begin VB.Label Label5 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Egresos"
         Height          =   285
         Left            =   45
         TabIndex        =   15
         Top             =   2520
         Width           =   1500
      End
      Begin VB.Label xTotIng 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   1545
         TabIndex        =   14
         Top             =   2220
         Width           =   1155
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Ingresos"
         Height          =   285
         Left            =   45
         TabIndex        =   13
         Top             =   2220
         Width           =   1500
      End
      Begin VB.Label xSumaCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   1545
         TabIndex        =   12
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label xBasico 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   285
         Left            =   1545
         TabIndex        =   11
         Top             =   1020
         Width           =   1155
      End
      Begin VB.Label xArea 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10.01 "
         Height          =   285
         Left            =   1545
         TabIndex        =   10
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label xCCosto 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10.02 "
         Height          =   285
         Left            =   1545
         TabIndex        =   9
         Top             =   1620
         Width           =   1155
      End
      Begin VB.Label Label13 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Fondo Pensiones"
         Height          =   285
         Left            =   45
         TabIndex        =   8
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label14 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Centro de Costo"
         Height          =   285
         Left            =   45
         TabIndex        =   7
         Top             =   1620
         Width           =   1500
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H80000010&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Totales por Trabajador"
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   45
         TabIndex        =   6
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Image ImgBuscar 
         Height          =   240
         Left            =   1620
         Picture         =   "InputPl.frx":098E
         Top             =   765
         Width           =   240
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Redondeo"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   45
         TabIndex        =   4
         Top             =   2820
         Width           =   1500
      End
      Begin VB.Label XREDONDEO 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   1545
         TabIndex        =   3
         Top             =   2820
         Width           =   1155
      End
      Begin VB.Label xBuscar 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Buscar "
         Height          =   285
         Left            =   1545
         TabIndex        =   5
         Top             =   720
         Width           =   1155
      End
   End
   Begin Crystal.CrystalReport Reporte 
      Left            =   3660
      Top             =   3105
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4275
      Left            =   45
      TabIndex        =   1
      Top             =   2505
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   7541
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Valor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Total General"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSDataGridLib.DataGrid dgInput 
      Height          =   2445
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   8685
      _ExtentX        =   15319
      _ExtentY        =   4313
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
      Caption         =   "Input de Datos de Planilla de Remuneraciones"
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
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7005
      Top             =   1770
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":0CD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":1024
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":1900
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":21DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":3000
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":3354
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "InputPl.frx":36A8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "InputPl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'TEMPORALES CON LOS QUE TRABAJA EL FORMULARIO
' [##TMPCNPADEL" & VGL_COMPUTER & "]
' [##TRABENVACA" & VGL_COMPUTER & "]
' [##PAGOSELIM" & VGL_COMPUTER & "]
' [##ADELELIM" & VGL_COMPUTER & "]

' [##ADELANTOS" & VGL_COMPUTER & "]
' [##TMP001" & VGL_COMPUTER & "]
' [##PAGOSCTACTE" & VGL_COMPUTER & "]
' [##TMPAUX" & VGL_COMPUTER & "]
' [##TMPCREPLAN" & VGL_COMPUTER & "]
' [##TMPQUINTA" & VGL_COMPUTER & "]
Dim WithEvents RSINPUT As ADODB.Recordset
Attribute RSINPUT.VB_VarHelpID = -1
Dim RSCNPT As New ADODB.Recordset
Dim RSAUX As New ADODB.Recordset
Dim XITEM As ListItem
Dim NUMV As Integer     'NUMERO DE CAMBIOS AFECTADOS
Dim STRCALCSUM As String, SWDELETE As Boolean
Dim SUMAVAR As String ' ACUMULA LOS CONCEPTOS VARIABLES DE QUINTA CATEGORIA

Private Sub Command1_Click()
    CambiaPanelBD True
        Call GRABAR_EN_INGMOV2000
        GRABARDATOS
    CambiaPanelBD False
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
Dim RSAUX As New ADODB.Recordset
RSAUX.Open "SELECT DISTINCT Codigo, Descripcion FROM  [##TMPCNPADEL" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
Set FrmDetAdel.DCmcampo.RowSource = RSAUX
   FrmDetAdel.CodigoTrab = RSINPUT!CODTRAB
   FrmDetAdel.DCmcampo.ListField = "DESCRIPCION"
   FrmDetAdel.DCmcampo.BoundColumn = "CODIGO"
   FrmDetAdel.DCmcampo.BoundText = "BASICO"
   FrmDetAdel.Caption = "Adelantos " & Me.Caption
   FrmDetAdel.Show 1
End Sub

Private Sub DGINPUT_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    RSINPUT.MOVE 0
End Sub

Private Sub dgInput_AfterUpdate()
   ' Call CALCULOTOTAL
End Sub

Private Sub DGINPUT_DblClick()
    If MsgBox("Desea ver datos del trabajador ", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    On Error GoTo FRMVIS
    VPTAREA = RSINPUT!CODTRAB
    If frValor.Visible Then Exit Sub
    Screen.MousePointer = 11
    CambiaPanelBD True
    Load frTrab
    CambiaPanelBD False
    frTrab.Show 1
    'EXIT SUB
    
    Set RSAUX = Nothing
    RSAUX.Open "SELECT CODTRAB, NOMBRES,AREA,CCOSTO,BASICO,ASIGFAM,FONDOPENS,TASA,APOROBLI,SEGURO,TOPESEGURO,COMISIONRA,UBIGEO,SEXO,TIPOTRAB,FECHAING,SITUACIÓN,CARGO,BANCO,ESSALUDVIDA,RUCEPS,NOPDT,OPCION01,OPCION02,OPCIONA,OPCIONB FROM TRABBOLETEAR WHERE CODTRAB='" & RSINPUT!CODTRAB & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    DBSTARPLAN.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET NOMBRES='" & RSAUX!NOMBRES & "',CODAREA='" & RSAUX!AREA & "',CODCCOSTO='" & RSAUX!CCosto & "',BASICO=" & RSAUX!BASICO & ",ASIGFAM=" & RSAUX!ASIGFAM & ",CODAFP='" & RSAUX!FONDOPENS & "',TASASCTR=" & RSAUX!TASA & ",APOROBL=" & RSAUX!APOROBLI & ",SEGURO=" & RSAUX!SEGURO & ",TOPESEGURO=" & RSAUX!TOPESEGURO & ",COMISIONRA=" & RSAUX!COMISIONRA & ",UBIGEO='" & RSAUX!UBIGEO & "',SEXO=" & RSAUX!Sexo & ",TIPOTRAB='" & RSAUX!TIPOTRAB & "',FECHAING=" & DateSQL(RSAUX!FECHAING) & ",SITUACION='" & RSAUX!SITUACIÓN & "',CARGO='" & RSAUX!CARGO & "',BANCO='" & RSAUX!BANCO & "',ESSALUDVIDA=" & IIf(RSAUX!ESSALUDVIDA, -1, 0) & ",RUCEPS='" & RSAUX!RUCEPS & "',NOPDT=" & RSAUX!NOPDT & ",OPCION01=" & RSAUX!OPCION01 & ",OPCION02=" & RSAUX!OPCION02 & ",OPCIONA='" & Trim(RSAUX!OPCIONA) & "',OPCIONB='" & Trim(RSAUX!OPCIONB) & "'  WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
    Set RSAUX = Nothing
    RSAUX.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    Dim OTROAUX As New ADODB.Recordset
    OTROAUX.Open "SELECT * FROM TRABAJADORES WHERE CODTRAB='" & RSINPUT!CODTRAB & "'", DBSYSTEM, adOpenStatic, adLockReadOnly
    Set RSAUX = Nothing
    RSAUX.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        Select Case RSAUX!TIPODATA
            Case "N"
                If Not IsNull(OTROAUX(Trim$(RSAUX!CODDATA))) Then
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & "=" & OTROAUX(Trim$(RSAUX!CODDATA)) & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                Else
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & "=0 WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                End If
            Case "T"
                If Not IsNull(OTROAUX(Trim$(RSAUX!CODDATA))) Then
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & " ='" & OTROAUX(Trim$(RSAUX!CODDATA)) & "' WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                Else
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & "='' WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                End If
            Case "F"
                If Not IsNull(OTROAUX(Trim$(RSAUX!CODDATA))) Then
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET " & RSAUX!CODDATA & " =" & DateSQL(OTROAUX(Trim$(RSAUX!CODDATA))) & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                Else
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET " & RSAUX!CODDATA & " =" & FechS("01/01/1900", Sqlf) & ", WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                End If
            Case "B"
                If OTROAUX(Trim$(RSAUX!CODDATA)) Then
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET " & RSAUX!CODDATA & "=-1 WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                Else
                    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET " & RSAUX!CODDATA & " =0 WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
                End If
        End Select
        RSAUX.MoveNext
    Loop
    Set Aux = Nothing
    RSINPUT!NOMBRES = RSINPUT!NOMBRES
    RSINPUT.Update
   ' Call CALCULOTOTAL(VPTAREA)
    RSINPUT.MOVE 0
    CambiaPanelBD False
    If frPersonal.hWnd <> 0 Then
        Unload frPersonal
    End If
    Exit Sub
FRMVIS:
    Resume Next
End Sub

Private Sub DGINPUT_HEADCLICK(ByVal COLINDEX As Integer)
    If COLINDEX > 1 Then
        Screen.MousePointer = 1
        If Not frValor.Visible Then
            Load frValor
            frValor.Show 1
            If VPTAREA = "0" Then Exit Sub
        End If
    Else
        CambiaPanelBD True
        RSINPUT.Sort = dgInput.Columns(COLINDEX).Caption
        CambiaPanelBD False
        Exit Sub
    End If
    SWDELETE = True
'    If VPTAREA <> "0" Then
        CambiaPanelBD True
        With RSINPUT
            .MoveFirst
            Do While Not .EOF
                .Fields(dgInput.Columns(COLINDEX).Caption).Value = Val(VPTAREA)
                .MoveNext
            Loop
            .MoveFirst
        End With
        CALCULOTOTAL
        SWDELETE = False
        RSINPUT.MoveFirst
        RSINPUT.MOVE 0
        CambiaPanelBD False
'    End If
End Sub

Private Sub DGINPUT_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    If SWDELETE Then Exit Sub
    If dgInput.COL = -1 Then Exit Sub
    Dim STRCOL As String
    RSCNPT.MoveFirst
    STRCOL = dgInput.Columns(dgInput.COL).Caption
    RSCNPT.FIND "CODIGO='" & STRCOL & "'"
    If RSCNPT.EOF Then Exit Sub
    xTotalCol.Caption = "TOTAL " & RSCNPT!NOMBRE
    Set RSAUX = Nothing
    RSAUX.Open "SELECT SUM(" & STRCOL & ") AS TOTALCOL FROM  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ", DBSYSTEM, adOpenStatic
    xSumaCol.Caption = Format(RSAUX!TOTALCOL, "0.00 ")
End Sub

Private Sub Form_Load()
    Dim XNV As Integer
    DBSYSTEM.Execute "UPDATE CONCEPTOS SET CRITERIO='' WHERE CRITERIO IS NULL"
    CambiaPanelBD True
    If Not ExisteCampo("REMUVAC", "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM) Then
        If ExisteTablaAux(" [##TRABENVACA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TRABENVACA" & VGL_COMPUTER & "]  "
            DBSYSTEM.Execute "SELECT CODTRAB INTO  [##TRABENVACA" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo.HISTOVAC WHERE (FECHAINI BETWEEN " & DateSQL(REGINPUT.FECHAINI) & " AND " & DateSQL(REGINPUT.FECHAFIN) & ") OR (FECHAFIN BETWEEN " & DateSQL(REGINPUT.FECHAINI) & " AND " & DateSQL(REGINPUT.FECHAFIN) & ") OR (" & DateSQL(REGINPUT.FECHAINI) & " BETWEEN FECHAINI AND FECHAFIN) OR (" & DateSQL(REGINPUT.FECHAFIN) & " BETWEEN FECHAINI AND FECHAFIN)", XNV
            DBSYSTEM.Execute "DELETE FROM  [##TRABENVACA" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM  [##TMPSELECT" & VGL_COMPUTER & "]) "
            DBSYSTEM.Execute "UPDATE  [##TRABENVACA" & VGL_COMPUTER & "]  SET CODTRAB=CODTRAB ", XNV
        If XNV > 0 Then
            If MsgBox("Existen " & XNV & " trabajadores en vacaciones durante este periodo, desea cargarlos de todas maneras", vbYesNo + vbQuestion) = vbNo Then
                DBSYSTEM.Execute "DELETE FROM  [##TMPSELECT" & VGL_COMPUTER & "]  WHERE CODTRAB IN (SELECT CODTRAB FROM  [##TRABENVACA" & VGL_COMPUTER & "]  ) "
            End If
        End If
    End If
    If GetSetting(App.CompanyName, "PLANILLAS", "Nando", "NO") <> "HOLA " Then On Error GoTo PasarError
    Screen.MousePointer = 11
    If ExisteTablaAux(" [##PAGOSELIM" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PAGOSELIM" & VGL_COMPUTER & "]  "
    DBSYSTEM.Execute "CREATE TABLE  [##PAGOSELIM" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CODMOV INT, CUOTA  Numeric(20,2) )"
         
    If ExisteTablaAux(" [##ADELELIM" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##ADELELIM" & VGL_COMPUTER & "]  "
    DBSYSTEM.Execute "CREATE TABLE  [##ADELELIM" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), CODIGO INT )"
        
    If ExisteTablaAux(" [##TMPCNPADEL" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCNPADEL" & VGL_COMPUTER & "]  "
    DBSYSTEM.Execute "Create Table  [##TMPCNPADEL" & VGL_COMPUTER & "]  (Codigo Varchar(20), Descripcion VarChar(100))"
    
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('BASICO', 'Remuneracion Basica')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('ASIGFAM', 'Asignacion familiar')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('TASASCTR', 'Tasa SCTR')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('APOROBL', 'Aportacion Obligatoria')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('SEGURO', 'Seguro')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('TOPESEGURO', 'Tope Seguro')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('COMISIONRA', 'Comision RA')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('TOTING', 'Total Ingreso')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('_HORAST', 'Horas Trabajadas')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('_HOREXTRAS', 'Horas Extras')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('_QUINTACAT', 'Quinta Categoria')"
    DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('BASICO', 'Remuneracion Basica')"
    
    Dim RsCAux2 As ADODB.Recordset

    If ExisteTablaAux("[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ") Then DBSYSTEM.Execute "DROP TABLE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]"
        STRCREA = "CREATE TABLE [##CALCINPUT" & Trim(VGL_COMPUTER) & "] (CODTRAB VARCHAR(8), NOMBRES VARCHAR(150), CODAREA VARCHAR(50), CODCCOSTO VARCHAR(50), BASICO  Numeric(20,2), ASIGFAM  Numeric(20,2), CODAFP VARCHAR(2), TASASCTR  Numeric(20,2), APOROBL  Numeric(20,2), SEGURO  Numeric(20,2), TOPESEGURO  Numeric(20,2), COMISIONRA  Numeric(20,2), SUMAAFP  Numeric(20,2), SUMASALUD  Numeric(20,2), TOTING  Numeric(20,2), TOTEGR  Numeric(20,2), _HORAST  Numeric(20,2), _HOREXTRAS  Numeric(20,2), _QUINTACAT  Numeric(20,2)"
        STRCREA = STRCREA + ", SUMAIES  Numeric(20,2), SUMARENTA  Numeric(20,2), SUMASCTR  Numeric(20,2), SUMACTS  Numeric(20,2), SUMAGRAT  Numeric(20,2), SUMAVAC  Numeric(20,2), T1  Numeric(20,2), T2  Numeric(20,2), T3  Numeric(20,2), T4  Numeric(20,2), T5  Numeric(20,2), OTROSING  Numeric(20,2), OTROSEGR  Numeric(20,2), ADELANTO  Numeric(20,2), UBIGEO VARCHAR(6), SEXO BIT, TIPOTRAB VARCHAR(2), FECHAING DATETIME, SITUACION VARCHAR(2), CARGO VARCHAR(150), BANCO VARCHAR(4), ESSALUDVIDA BIT, RUCEPS VARCHAR(11), NOPDT BIT, OPCION01 BIT, OPCION02 BIT, OPCIONA VARCHAR(15), OPCIONB VARCHAR(15), XREDONDEO  Numeric(20,2), AFECTOQUINTA BIT NOT NULL DEFAULT 0)"
        DBSYSTEM.Execute STRCREA
    
    'SE SELECCIONA LOS CONCEPTOS CON DEL FORMATO ESPECIFICADO
    RSCNPT.Open "SELECT CONCEPTOS.* FROM CONCEPTOS,FORMARUBS WHERE CONCEPTO=CODIGO AND ID_FORMATO=" & CalcPlan.xFormato.Tag & " ORDER BY TIPO, FILA", DBSYSTEM, adOpenStatic
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
'            If RSCNPT!CODIGO = "DIASTRAB" Then Stop
            DBSYSTEM.Execute "ALTER TABLE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ADD " & RSCNPT!Codigo & "  Numeric(20,2)"
                Set RsCAux2 = New ADODB.Recordset
                RsCAux2.Open "SELECT CODIGO, NOMBRE FROM CONCEPTOS WHERE CODIGO ='" & RSCNPT!Codigo & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
                    If RsCAux2.RecordCount > 0 Then
                        DBSYSTEM.Execute "insert into  [##TMPCNPADEL" & VGL_COMPUTER & "] (Codigo,Descripcion) VALUES('" & RSCNPT!Codigo & "', '" & RsCAux2!NOMBRE & "')"
                    End If
            .MoveNext
        Loop
    End With
    DBSYSTEM.Execute "INSERT INTO  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   (CODTRAB,NOMBRES,CODAREA,CODCCOSTO,BASICO,ASIGFAM,CODAFP,TASASCTR,APOROBL,SEGURO,TOPESEGURO,COMISIONRA,UBIGEO,SEXO,TIPOTRAB,FECHAING,SITUACION,CARGO,BANCO,ESSALUDVIDA,RUCEPS,NOPDT,OPCION01,OPCION02,OPCIONA,OPCIONB,XREDONDEO,AFECTOQUINTA) SELECT CODTRAB, NOMBRES,AREA,CCOSTO,BASICO,ASIGFAM,FONDOPENS,TASA,APOROBLI,SEGURO,TOPESEGURO,COMISIONRA,UBIGEO,SEXO,TIPOTRAB,FECHAING,SITUACIÓN,CARGO,BANCO,ESSALUDVIDA,RUCEPS,NOPDT,OPCION01,OPCION02,OPCIONA,OPCIONB,XREDONDEO,AFECTOQUINTA FROM " & REGSISTEMA.BASESQL & ".dbo.TRABBOLETEAR WHERE CODTRAB IN (SELECT CODTRAB FROM  [##TMPSELECT" & VGL_COMPUTER & "] )" & REGINPUT.CADENA & " ORDER BY NOMBRES", X
    'ASIGNAR CERO A TODAS LAS COLUMNAS
    'SI NO SE PONEN A CERO NO SE EJECUTAN LAS SUMAS POR BLOQUES, PUESTO QUE ESTA EN NULL
    Set RSAUX = Nothing
    'ASIGNACIÓN DEL PRECIO DEL DOLAR
    DBSYSTEM.Execute "ALTER TABLE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   ADD VALORDOLAR  Numeric(20,2)"
    DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET VALORDOLAR=" & MDIPrincipal.BarraEstado.Panels("Dolar").Text
    If Not ExisteTabla("DATATRAB") Then
        DBSYSTEM.Execute "CREATE TABLE DATATRAB (CODDATA VARCHAR(15), DESCDATA VARCHAR(30), TIPODATA VARCHAR(1))"
        MsgBox "El Sistema de PLANILLAS ha actualizado su versión", vbInformation
    End If
    RSAUX.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockReadOnly
    Do While Not RSAUX.EOF
        Select Case RSAUX!TIPODATA
            Case "N"
                DBSYSTEM.Execute "ALTER TABLE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   ADD " & RSAUX!CODDATA & "  Numeric(20,2)"
                DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ." & RSAUX!CODDATA & "=B." & RSAUX!CODDATA & " FROM  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   A, [" & REGSISTEMA.BASESQL & "].dbo.TRABAJADORES B WHERE A.CODTRAB=B.CODTRAB"
                DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET " & RSAUX!CODDATA & " =0 WHERE " & RSAUX!CODDATA & " IS NULL"
            Case "T"
                DBSYSTEM.Execute "ALTER TABLE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   ADD " & RSAUX!CODDATA & " VARCHAR(30)"
                DBSYSTEM.Execute "UPDATE  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   SET  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ." & RSAUX!CODDATA & "=B." & RSAUX!CODDATA & " FROM  [##CALCINPUT" & Trim(VGL_COMPUTER) & "]   A, [" & REGSISTEMA.BASESQL & "].dbo.TRABAJADORES B WHERE A.CODTRAB=B.CODTRAB"
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & " ='' WHERE " & RSAUX!CODDATA & " IS NULL"
            Case "F"
                DBSYSTEM.Execute "ALTER TABLE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ADD " & RSAUX!CODDATA & " DATETIME"
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ." & RSAUX!CODDATA & "=B." & RSAUX!CODDATA & " FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  A, [" & REGSISTEMA.BASESQL & "].dbo.TRABAJADORES B WHERE A.CODTRAB=B.CODTRAB"
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & " ='01/01/1900' WHERE " & RSAUX!CODDATA & " IS NULL"
            Case "B"
                DBSYSTEM.Execute "ALTER TABLE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  ADD " & RSAUX!CODDATA & " BIT"
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ." & RSAUX!CODDATA & "=B." & RSAUX!CODDATA & " FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  A, [" & REGSISTEMA.BASESQL & "].dbo.TRABAJADORES B WHERE A.CODTRAB=B.CODTRAB"
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CODDATA & " =0 WHERE " & RSAUX!CODDATA & " IS NULL"
            Case Else
                MsgBox "ERROR DE USUARIO: La Configuración de un campo agregado a la Base de Datos de trabajadores no es valido ", vbCritical
        End Select
        RSAUX.MoveNext
    Loop
    Set RSAUX = Nothing
    DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET SUMAAFP=0,SUMASALUD=0,TOTING=0,TOTEGR=0,_HORAST=0,_HOREXTRAS=0,_QUINTACAT=0,SUMAIES=0,SUMARENTA=0,SUMASCTR=0,SUMACTS=0,SUMAGRAT=0,SUMAVAC=0,T1=0,T2=0,T3=0,T4=0,T5=0,OTROSING=0,OTROSEGR=0", X
    With RSCNPT
        .MoveFirst
        STRCREA = "ADELANTO=0"
        Do While Not .EOF
            STRCREA = STRCREA & ", " & RSCNPT!Codigo & "=0"
            .MoveNext
        Loop
    End With
    DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & STRCREA
    '-FIN DE LA ACTUALIZACIÓN A CERO
    
    RSCNPT.MoveFirst
    'CARGA DE LOS DATOS TIPO FORMULA
    Do While Not RSCNPT.EOF
        If Not RSCNPT!ESESCRITO Then
            If RSCNPT!FORMULA = "" Then
                MsgBox "El ingreso de la Fórmula no ha sido registrada, el Sistema generará constantes avisos de errores sobre el fallo del siguiente rubro: " & RSCNPT!Codigo & ": " & RSCNPT!NOMBRE
                If MsgBox("Desea continuar con la carga del Sistema ", vbYesNo) = vbNo Then
                    Set RSCNPT = Nothing
                    Unload Me
                End If
            Else
                'OJO QUE LOS TIPOS INFORMATIVOS NO PUEDEN TENER FORMULAS
                Set XITEM = Lista.ListItems.Add(, RSCNPT!Codigo, RSCNPT!Codigo, , RSCNPT!TIPO + 1)
                XITEM.SubItems(1) = RSCNPT!NOMBRE
                XITEM.SubItems(2) = "0.00"
                XITEM.Tag = 1
            End If
        End If
        RSCNPT.MoveNext
    Loop
    Lista.ColumnHeaders(4).Width = 1154.835
    Lista.ColumnHeaders(3).Width = 959.8111
    Lista.ColumnHeaders(2).Width = 2355.024
    
    'CARGA LAS FORMULAS DE INICIO DEL PROCESO
    'FORMULAS COMO SUMAAFP, TOTAL INGRESOS, ENTRE OTRAS
    Dim CADSUMAS(14) As String
    CADSUMAS(0) = "0+round(OTROSING,2)"
    CADSUMAS(1) = "0+round(OTROSING,2)"
    CADSUMAS(2) = "0+round(OTROSING,2)"
    CADSUMAS(3) = "0+round(OTROSING,2)"
    CADSUMAS(4) = "0+round(OTROSING,2)"
    CADSUMAS(5) = "0+round(OTROSING,2)"
    CADSUMAS(6) = "0+round(OTROSING,2)"
    CADSUMAS(7) = "0+round(OTROSING,2)"
    CADSUMAS(8) = "0+round(OTROSING,2)"
    CADSUMAS(9) = "0+round(OTROSING,2)"
    CADSUMAS(10) = "0+round(OTROSING,2)"
    CADSUMAS(11) = "0+round(OTROSING,2)"
    CADSUMAS(12) = "0+round(OTROSING,2)"
    CADSUMAS(13) = "0+round(OTROSING,2)"
    Dim CADHORAS As String, CADHOREXT As String, CAD5TA As String
    CADHORAS = "0"
    CADHOREXT = "0"
    CAD5TA = "0"
    'RECALCULAR LOS TOTALES DE INGRESOS AFECTOS A ...
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO = 0 Then
                'EN ESTA SECCIÓN SE CALCULAN LOS TIPOS DE INFORMACIÓN
                'COMO HORAS TRABAJADAS Y HORAS EXTRAS
                If !TIPOINFO < 3 Then
                    Select Case !TIPOINFO
                        Case 0: CADHORAS = CADHORAS & "+" & !Codigo & "* 8"
                        Case 1: CADHORAS = CADHORAS & "+" & !Codigo
                        Case 2: CADHOREXT = CADHOREXT & "+" & !Codigo
                    End Select
                End If
            End If
            If !TIPO = 1 Then
                CADSUMAS(0) = CADSUMAS(0) & IIf(CADSUMAS(0) = "", "", "+") & !Codigo  'SOLO PARA EL CÁLCULO DE INGRESOS
                If !SUMAAFP Then CADSUMAS(1) = CADSUMAS(1) & IIf(CADSUMAS(1) = "", "", "+") & !Codigo
                If !SUMASALUD Then CADSUMAS(2) = CADSUMAS(2) & IIf(CADSUMAS(2) = "", "", "+") & !Codigo
                If !SUMAIES Then CADSUMAS(3) = CADSUMAS(3) & IIf(CADSUMAS(3) = "", "", "+") & !Codigo
                If !SUMARENTA Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(4) = CADSUMAS(4) & IIf(CADSUMAS(4) = "", "", "+") & Trim(!COMENTARIO)
                        Else
                            CADSUMAS(4) = CADSUMAS(4) & IIf(CADSUMAS(4) = "", "", "+") & !Codigo
                        End If
                End If
                If !SUMASCTR Then CADSUMAS(5) = CADSUMAS(5) & IIf(CADSUMAS(5) = "", "", "+") & !Codigo
                If !SUMACTS Then CADSUMAS(6) = CADSUMAS(6) & IIf(CADSUMAS(6) = "", "", "+") & !Codigo
                If !SUMAGRAT Then CADSUMAS(7) = CADSUMAS(7) & IIf(CADSUMAS(7) = "", "", "+") & !Codigo
                If !SUMAVAC Then CADSUMAS(8) = CADSUMAS(8) & IIf(CADSUMAS(8) = "", "", "+") & !Codigo
                If !SUMAT1 Then CADSUMAS(9) = CADSUMAS(9) & IIf(CADSUMAS(9) = "", "", "+") & !Codigo
                If !SUMAT2 Then CADSUMAS(10) = CADSUMAS(10) & IIf(CADSUMAS(10) = "", "", "+") & !Codigo
                If !SUMAT3 Then CADSUMAS(11) = CADSUMAS(11) & IIf(CADSUMAS(11) = "", "", "+") & !Codigo
                If !SUMAT4 Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(12) = CADSUMAS(12) & IIf(CADSUMAS(12) = "", "", "+") & !COMENTARIO
                        Else
                            CADSUMAS(12) = CADSUMAS(12) & IIf(CADSUMAS(12) = "", "", "+") & !Codigo
                        End If
                End If
                If !SUMAT5 Then
                        If Len(Trim(!COMENTARIO)) Then
                            CADSUMAS(13) = CADSUMAS(13) & IIf(CADSUMAS(13) = "", "", "+") & !COMENTARIO
                        Else
                            CADSUMAS(13) = CADSUMAS(13) & IIf(CADSUMAS(13) = "", "", "+") & !Codigo
                            
                            
                        End If
                End If
            End If
            .MoveNext
        Loop
    End With
    SUMAVAR = CADSUMAS(13)
    STRCALCSUM = "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET  TOTING=" & CADSUMAS(0) & ", SUMAAFP=" & CADSUMAS(1) & ", SUMASALUD=" & CADSUMAS(2) & ", SUMAIES=" & CADSUMAS(3) & ", SUMARENTA=" & CADSUMAS(4) & ", SUMASCTR=" & CADSUMAS(5) & ", SUMACTS=" & CADSUMAS(6) & ", SUMAGRAT=" & CADSUMAS(7) & ", SUMAVAC=" & CADSUMAS(8) & ", T1=" & CADSUMAS(9) & ", T2=" & CADSUMAS(10) & ", T3=" & CADSUMAS(11) & ", T4=" & CADSUMAS(12) & ", T5=" & CADSUMAS(13) & ", _HORAST=" & CADHORAS & ", _HOREXTRAS=" & CADHOREXT & ", _QUINTACAT=" & CAD5TA
    
    'VACIADO DE DATOS DE ASISTENCIA DE TRABAJADORES
    Dim AUXCAD As String
    If CalcPlan.Check2.Value = 1 Then 'SI CARGAR ASISTENCIA
        Set RSAUX = Nothing
        RSAUX.Open "SELECT CODTRAB, CONCEPTO, SUM(VALOR) AS CANTI FROM ASIS" & REGSISTEMA.ANNO & " WHERE (DIA BETWEEN " & DateSQL(REGINPUT.FECHAINI) & " AND " & DateSQL(REGINPUT.FECHAFIN) & ") AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )  GROUP BY CODTRAB, CONCEPTO", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            'Validando si el campo actualizar existe en el formato
            If Not ExisteCampo(RSAUX!CONCEPTO, "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM) Then
                 If RSAUX!CONCEPTO <> AUXCAD Then
                    MsgBox "El Campo """ & RSAUX!CONCEPTO & """ no existe en el formato de Planilla utilizado " & Chr(13) & "Por lo tanto no aparecera en la planilla ", vbInformation
                    AUXCAD = RSAUX!CONCEPTO
                 End If
               Else
                 DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CONCEPTO & "=" & RSAUX!CANTI & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            End If
            RSAUX.MoveNext
        Loop
    End If
    
    AUXCAD = ""
    'VACIADO DE DATOS DE MOVIMIENTOS
    If CalcPlan.Check3.Value = 1 Then
        Set RSAUX = Nothing
        RSAUX.Open "SELECT CODTRAB, CONCEPTO, SUM(VALOR) AS CANTI FROM INGMOV2000 WHERE CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ) AND CODNOMBOL=" & REGINPUT.Codigo & " GROUP BY CODTRAB, CONCEPTO", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            If Not ExisteCampo(RSAUX!CONCEPTO, "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM) Then
                If RSAUX!CONCEPTO <> AUXCAD Then
                    MsgBox "El Campo """ & RSAUX!CONCEPTO & """ no existe en el formato de Planilla utilizado " & Chr(13) & "Por lo tanto no aparecera en la planilla ", vbInformation
                    AUXCAD = RSAUX!CONCEPTO
                End If
              Else
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CONCEPTO & "=" & RSAUX!CANTI & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            End If
            RSAUX.MoveNext
        Loop
    End If
    
    'CAPTURA DE LOS ADELANTOS DE VACACIONES
    Dim VarAdelVac As String
    'Configuracion de donde cargaREGINPUT.CODIGO
    VarAdelVac = DevuelveValor("SELECT ADELVAC FROM EMPRESA", DBSYSTEM)
    If IsNull(VarAdelVac) Then VarAdelVac = ""
    
    If CalcPlan.Check7.Value = 1 And ExisteCampo(VarAdelVac, "CALCINPUT", DBSYSTEM) Then
        Set RSAUX = Nothing
        RSAUX.Open "SELECT (TOTING-TOTEGR) AS NETO, CODTRAB FROM " & REGINPUT.BOL_TABLE & " WHERE CODNOMBOL=" & REGINPUT.Codigo & " AND CODTRAB In (SELECT CODTRAB FROM HISTOVAC WHERE NOMBOL=" & REGINPUT.Codigo & " AND CERRADO=1)", DBSYSTEM, adOpenStatic, adLockReadOnly
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & VarAdelVac & "=" & RSAUX!Neto & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            RSAUX.MoveNext
        Loop
    End If
    

    'CAPTURA DE LOS TRABAJADORES QUE YA TIENEN BOLETAS DE REMUNERACIONES
    Dim XNUMANTERIORES As Long
    If ExisteTablaAux(" [##BOLSANT" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##BOLSANT" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "SELECT CODTRAB, INUMBOL, FECHA, TOTING, TOTEGR INTO  [##BOLSANT" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo." & REGINPUT.BOL_TABLE & " WHERE CODNOMBOL=" & REGINPUT.Codigo & " AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )", XNUMANTERIORES
    
    'CARGAR LOS DATOS YA INGRESADOS EN BOLETAS, PARA DESPUES NO ESTAR RETIPIANDO LOS VALORES
    If XNUMANTERIORES > 0 Then
        Screen.MousePointer = 1
        If MsgBox("Desea cargar los datos ingresados anteriormente si escoge SI" & vbCrLf & " cargara la información del las  Boletas,si escohe NO " & vbCrLf & "cargara las ulimas modificaciones en Base Datos. ", vbYesNo + vbQuestion) = vbYes Then
            Screen.MousePointer = 11
            Set RSAUX = Nothing
            'RSAUX.Open "SELECT CONCEPTO, MONTO, BOL.INUMBOL, CODTRAB FROM " & REGINPUT.MOV_TABLE & " MOV," & REGINPUT.BOL_TABLE & " BOL WHERE BOL.INUMBOL=MOV.INUMBOL AND BOL.INUMBOL IN (SELECT  [##BOLSANT" & VGL_COMPUTER & "] .INUMBOL FROM STARPLAN.dbo. [##BOLSANT" & VGL_COMPUTER & "] )AND MOV.CONCEPTO IN (SELECT CONCEPTOS.CODIGO FROM CONCEPTOS WHERE ESESCRITO=1)", DBSYSTEM, adOpenStatic, adLockReadOnly
            'MODIFICADO LISTO  PARA HABILITAR*****************
            RSAUX.Open "SELECT CONCEPTO, MONTO, BOL.INUMBOL, CODTRAB FROM " & REGINPUT.MOV_TABLE & " MOV," & REGINPUT.BOL_TABLE & " BOL WHERE BOL.INUMBOL=MOV.INUMBOL AND BOL.INUMBOL IN (SELECT  [##BOLSANT" & VGL_COMPUTER & "].INUMBOL FROM [##BOLSANT" & VGL_COMPUTER & "] )AND MOV.CONCEPTO IN (SELECT CONCEPTOS.CODIGO FROM CONCEPTOS WHERE ESESCRITO=1)", DBSYSTEM, adOpenStatic, adLockReadOnly
            '**************************************************
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & RSAUX!CONCEPTO & "=" & RSAUX!MONTO & "  WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
                RSAUX.MoveNext
            Loop
            'modficacion para llenar el bASICO
            Set RSAUX = Nothing
            RSAUX.Open "SELECT DISTINCT CODTRAB,BOL.BASICO FROM " & REGINPUT.MOV_TABLE & " MOV," & REGINPUT.BOL_TABLE & " BOL WHERE BOL.INUMBOL=MOV.INUMBOL AND BOL.INUMBOL IN (SELECT  [##BOLSANT" & VGL_COMPUTER & "].INUMBOL FROM [##BOLSANT" & VGL_COMPUTER & "] )", DBSYSTEM, adOpenStatic, adLockReadOnly
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET BASICO=" & RSAUX!BASICO & "  WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
                RSAUX.MoveNext
            Loop
            'modficacion para llenar el bASICO
        End If
        Screen.MousePointer = 11
    End If
    
    'VACIADO DE DATOS DE ADELANTOS DE REMUNERACIONES
    '*CAMBIO
    If CalcPlan.Check1.Value = 1 Then 'TABLA DE ADELANTOS0
        If ExisteTablaAux(" [##ADELANTOS" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##ADELANTOS" & VGL_COMPUTER & "] "
        If CalcPlan.Check5 Then
            DBSYSTEM.Execute "SELECT CODIGO,CODTRAB,MES,FECHAING,MONTO, NOMBOL INTO  [##ADELANTOS" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.BASESQL & ".dbo." & REGSISTEMA.TABLAADEL & "  ORDER BY CODTRAB"
        Else
            'DBSYSTEM.Execute "INSERT INTO  [##ADELANTOS" & VGL_COMPUTER & "]  SELECT CODIGO, CODTRAB, MES, FECHAING, MONTO,NOMBOL FROM " & REGSISTEMA.BASESQL & ".dbo." & REGSISTEMA.TABLAADEL & " WHERE ORIGEN=" & REGINPUT.CODIGO & " AND CODTRAB NOT IN (SELECT CODTRAB FROM  [##BOLSANT" & VGL_COMPUTER & "] ) ORDER BY CODTRAB"
        End If
        DBSYSTEM.Execute "CREATE TABLE  [##ADELANTOS" & VGL_COMPUTER & "]  (CODIGO INT, CODTRAB VARCHAR(8), MES DATETIME, FECHAING DATETIME, MONTO  Numeric(20,2), NOMBOL INT)"
        DBSYSTEM.Execute "DELETE FROM  [##ADELANTOS" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )"
        'SE AGREGAN TAMBIÉN LOS ADELANTOS YA COBRADOS. MODO EDICION, PASAN CON NOMBOL
        DBSYSTEM.Execute "INSERT INTO  [##ADELANTOS" & VGL_COMPUTER & "]  (CODTRAB,MES,FECHAING,MONTO,NOMBOL) SELECT CODTRAB, MES, FECHAING, MONTO, NOMBOL FROM " & REGSISTEMA.BASESQL & ".dbo.ADEL" & REGSISTEMA.ANNO & " WHERE ORIGEN=" & REGINPUT.Codigo & " AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )"
            
            'FECHA DE MODIFICACION 24/09/2001---NORKA NORKA
            'NUEVOS CAMBIOS
            '##DETALLEADEL
            'DBSYSTEM.Execute "INSERT INTO  [##ADELANTOS" & VGL_COMPUTER & "]  (CODTRAB,MES,FECHAING,MONTO,NOMBOL) SELECT CODTRAB, MES, FECHAING, MONTO, NOMBOL FROM ##DETALLEADEL"
            'FIN DE LA MODIFICACION 24/09/2001---NORKA NORKA
            
        Set RSAUX = Nothing
        RSAUX.Open " [##ADELANTOS" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET ADELANTO=ADELANTO+" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            RSAUX.MoveNext
        Loop
    End If
    
    'VACIADO DE DATOS DE CUENTAS CORRIENTES
    Dim CADPAG As String
    If CalcPlan.Check4.Value = 1 Then
        If ExisteTablaAux(" [##TMP001" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMP001" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "UPDATE MOVICTA SET CUOTA=0 WHERE PROGRAMADO=1"
        'LE DESCUENTO EL PRESTAMO PROGRAMADO PARA ESTA FECHA
        DBSYSTEM.Execute "SELECT * INTO  [##TMP001" & VGL_COMPUTER & "]  FROM CTACTEPROG  C " & _
        " WHERE FECHA BETWEEN " & FechS(REGINPUT.FECHAINI, Sqlf) & " AND " & FechS(REGINPUT.FECHAFIN, Sqlf) & " AND " & _
        " ISNULL((SELECT P.TIPOBOLETA " & _
        " FROM PAGOSCTA P WHERE P.CODMOV=C.CODMOV  AND P.CODNOMBOL=" & REGINPUT.Codigo & " AND P.TIPOBOLETA='A' AND P.SECUENCIA=C.SECUENCIA),'')<>'A' "
        
        DBSYSTEM.Execute "UPDATE MOVICTA SET MOVICTA.CUOTA=B.IMPORTE,MOVICTA.ULTSECU=B.SECUENCIA FROM MOVICTA A, [##TMP001" & VGL_COMPUTER & "]  B WHERE A.CODMOV=B.CODMOV "
        DBSYSTEM.Execute "UPDATE MOVICTA SET CUOTA=CUOTA WHERE PROGRAMADO=1"
        If ExisteTablaAux(" [##PAGOSCTACTE" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##PAGOSCTACTE" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "CREATE TABLE  [##PAGOSCTACTE" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8) NULL, CODMOV INT NULL, DESCRIPCION VARCHAR(50) NULL, TIPOGRUPO INT NULL, CUOTA  Numeric(20,2) NULL, MONEDA INT NULL,ULTSECU INT NULL)"
        If CalcPlan.Check6.Value Then
            'DBSYSTEM.Execute "INSERT INTO  [##PAGOSCTACTE" & VGL_COMPUTER & "]  SELECT CODTRAB, CODMOV, DESCRIPCION, TIPOGRUPO, CUOTA=CASE WHEN MOVICTA.SALDO > MOVICTA.CUOTA THEN MOVICTA.CUOTA*((100- PORCQUINC)/100) ELSE MOVICTA.SALDO END, MONEDA,ULTSECU FROM " & REGSISTEMA.BASESQL & ".dbo.MOVICTA WHERE SALDO>0 AND FECHAINI <=" & DateSQL(REGINPUT.FECHAFIN)
            CADPAG = "INSERT INTO  [##PAGOSCTACTE" & VGL_COMPUTER & "] SELECT CODTRAB, CODMOV, DESCRIPCION, TIPOGRUPO, " & _
                     "CASE WHEN (SELECT TOP 1  ISNULL(P.TIPOBOLETA,'') FROM PAGOSCTA P   WHERE P.CODMOV=MOVICTA.CODMOV AND  " & _
                     "P.CODNOMBOL=" & REGINPUT.Codigo & ") <>'' THEN " & _
                     "CASE WHEN MOVICTA.CUOTA >SALDO THEN SALDO ELSE MOVICTA.CUOTA * ((100- ISNULL(PORCQUINC,0))/100) END ELSE  " & _
                     "CASE WHEN MOVICTA.CUOTA >SALDO THEN SALDO ELSE MOVICTA.CUOTA END END  AS  CUOTA, " & _
                     "MONEDA,ULTSECU FROM MOVICTA WHERE SALDO>0.001  AND FECHAINI <=" & FechS(REGINPUT.FECHAFIN, Sqlf)
        Else
            'DBSYSTEM.Execute "INSERT INTO  [##PAGOSCTACTE" & VGL_COMPUTER & "]  SELECT CODTRAB, CODMOV, DESCRIPCION, TIPOGRUPO, CUOTA=CASE WHEN MOVICTA.SALDO > MOVICTA.CUOTA THEN MOVICTA.CUOTA*((100- PORCQUINC)/100) ELSE MOVICTA.SALDO END, MONEDA,ULTSECU FROM " & REGSISTEMA.BASESQL & ".dbo.MOVICTA WHERE SALDO>0  AND FECHAINI <=" & DateSQL(REGINPUT.FECHAFIN) & " AND CODTRAB NOT IN (SELECT CODTRAB FROM  [##BOLSANT" & VGL_COMPUTER & "] )"
            CADPAG = "INSERT INTO  [##PAGOSCTACTE" & VGL_COMPUTER & "] SELECT CODTRAB, CODMOV, DESCRIPCION, TIPOGRUPO, " & _
                     "CASE WHEN (SELECT TOP 1  ISNULL(P.TIPOBOLETA,'') FROM PAGOSCTA P   WHERE P.CODMOV=MOVICTA.CODMOV AND  " & _
                     "P.CODNOMBOL=" & REGINPUT.Codigo & ") <>'' THEN " & _
                     "CASE WHEN MOVICTA.CUOTA >SALDO THEN SALDO ELSE MOVICTA.CUOTA * ((100- ISNULL(PORCQUINC,0))/100) END ELSE  " & _
                     "CASE WHEN MOVICTA.CUOTA >SALDO THEN SALDO ELSE MOVICTA.CUOTA END END  AS  CUOTA, " & _
                     "MONEDA,ULTSECU FROM MOVICTA WHERE SALDO>0.001  AND FECHAINI <=" & FechS(REGINPUT.FECHAFIN, Sqlf) & " AND CODTRAB NOT IN (SELECT CODTRAB FROM [##BOLSANT" & VGL_COMPUTER & "])"
        End If
 DBSYSTEM.Execute CADPAG
        DBSYSTEM.Execute "DELETE FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODTRAB NOT IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )"
        DBSYSTEM.Execute "DELETE FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODTRAB IS NULL"
        'SE CARGAN TAMBIÉN LOS PAGOS YÁ REALIZADOS. LOS YA COBRADOS
        DBSYSTEM.Execute "ALTER TABLE  [##PAGOSCTACTE" & VGL_COMPUTER & "]  ADD CODNOMBOL INT"
        DBSYSTEM.Execute "UPDATE  [##PAGOSCTACTE" & VGL_COMPUTER & "]  SET CODNOMBOL=0"
        'USAMOS UN TEMPORAL EN EL DBSYSTEM PARA LUEGO TRASPASARLO A LA BASE TEMPORAL DEL CLIENTE
        If ExisteTablaAux(" [##TMPAUX" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPAUX" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "SELECT PAGOSCTA.CODTRAB, MOVICTA.CODMOV, DESCRIPCION, TIPOGRUPO, MONTO AS CUOTA, MONEDA,SECUENCIA AS ULTSECU, CODNOMBOL INTO  [##TMPAUX" & VGL_COMPUTER & "]  FROM MOVICTA, PAGOSCTA WHERE MOVICTA.CODMOV=PAGOSCTA.CODMOV AND TIPOBOLETA='B' AND PAGOSCTA.CODNOMBOL=" & REGINPUT.Codigo
        Set RSAUX = Nothing
        RSAUX.Open " [##TMPAUX" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
        
        'ESPERA FORZADA POR ADO
'        Load frWait
'        frWait.Timer1.Interval = 300
'        frWait.Show 1
        DBSYSTEM.Execute "INSERT INTO  [##PAGOSCTACTE" & VGL_COMPUTER & "]  SELECT * FROM  [##TMPAUX" & VGL_COMPUTER & "] "
        DBSYSTEM.Execute "DROP TABLE  [##TMPAUX" & VGL_COMPUTER & "]  "
        Set RSAUX = Nothing
        RSAUX.Open "SELECT CODTRAB, TIPOGRUPO, SUM(CUOTA) AS MONTO FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  GROUP BY CODTRAB, TIPOGRUPO", DBSYSTEM, adOpenStatic
        Do While Not RSAUX.EOF
            If RSAUX!TIPOGRUPO = 1 Then
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSING=" & ESNULO(RSAUX!MONTO, 0) & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            Else
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSEGR=" & ESNULO(RSAUX!MONTO, 0) & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
            End If
            RSAUX.MoveNext
        Loop
    End If
    
    '------------------------------------------
    'JALA EL CÁLCULO DE VACACIONES
    '------------------------------------------
    Set RSAUX = Nothing
    RSCNPT.MoveFirst
    RSCNPT.FIND "CODIGO='REMUVAC'"
    If Not RSCNPT.EOF Then
        RSAUX.Open "SELECT CODTRAB, MONTO FROM HISTOVAC WHERE NOMBOL=" & REGINPUT.Codigo & " AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )", DBSYSTEM, adOpenStatic
        If RSAUX.RecordCount > 0 Then
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET REMUVAC=" & RSAUX!MONTO & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
                RSAUX.MoveNext
            Loop
        End If
    End If
    
    '------------------------------------------
    'JALA EL CÁLCULO DE GRATIFICACIONES
    '------------------------------------------
    Set RSAUX = Nothing
    RSCNPT.MoveFirst
    RSCNPT.FIND "CODIGO='REMUGRAT'"
    If Not RSCNPT.EOF Then
        RSAUX.Open "SELECT PLANGRATI.* FROM PLANGRATI, GRATIFICACION WHERE PLANGRATI.CODIGO=GRATIFICACION.CODIGO AND GRATIFICACION.PERIODO=" & REGINPUT.Codigo & " AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  )", DBSYSTEM, adOpenStatic
        If RSAUX.RecordCount > 0 Then
            Do While Not RSAUX.EOF
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET REMUGRAT=" & RSAUX!IMPORTEGRATI & " WHERE CODTRAB='" & RSAUX!CODTRAB & "'"
                RSAUX.MoveNext
            Loop
        End If
    End If
    
    'LLAMADA A LOS CALCULOS TOTALES Y QUE SERAN COLOCADOS EN EL CALCINPUT
    SWDELETE = False
    '------------------------------------------
    '------------------------------------------
    CALCULOTOTAL
    '------------------------------------------
    '------------------------------------------
    
    'APERTURA DE LA TABLA DE CALCINPUT, DONDE SE ESCRIBIRÁN LOS DATOS Y
    'SE CALCULARÁN LAS FORMULAS. TAMBIÉN ALLÍ SE COLOCAN LAS SUMAS
    With RSCNPT
        .MoveFirst
        STRCREA = "CODTRAB, NOMBRES"
        Do While Not .EOF
            If RSCNPT!ESESCRITO Then STRCREA = STRCREA & ", " & RSCNPT!Codigo
            .MoveNext
        Loop
    End With
    Set RSINPUT = New ADODB.Recordset
    RSINPUT.Open "SELECT " & STRCREA & " FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenDynamic, adLockOptimistic
    Set dgInput.DataSource = RSINPUT
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If RSCNPT!ESESCRITO Then
                If RSCNPT!FLAG = 1 Then
                    dgInput.Columns(RSCNPT!Codigo).Locked = True
                End If
                dgInput.Columns(RSCNPT!Codigo).Alignment = dbgRight
                dgInput.Columns(RSCNPT!Codigo).NumberFormat = "0.00 "
                dgInput.Columns(RSCNPT!Codigo).Width = 794.8347
            End If
            .MoveNext
        Loop
    End With
    dgInput.Columns("CODTRAB").Locked = True
    dgInput.Columns("NOMBRES").Locked = True
    xNumTrabs.Caption = " " & RSINPUT.RecordCount & " TRABAJADORES"
    Me.Caption = "PLANILLA DE REM.: " & CalcPlan.Lista.SelectedItem.Text
    If Not ExisteCampo("XREDONDEO", REGINPUT.BOL_TABLE, DBSYSTEM) Then
        'COMPATIBILIDAD CON VERSIONES ANTERIORES
        DBSYSTEM.Execute "ALTER TABLE " & REGINPUT.BOL_TABLE & " ADD XREDONDEO  Numeric(20,2)"
        DBSYSTEM.Execute "UPDATE " & REGINPUT.BOL_TABLE & " SET XREDONDEO=0"
    End If
Screen.MousePointer = 1
    'Call CALCULOXREGISTRO(RSINPUT)
    Call CALCULOTOTAL
    Call CALCULOXREGISTRO(RSINPUT)
    
    CambiaPanelBD False
    Exit Sub
PasarError:
    Open App.PATH & "\ERRMSG.TXT" For Append As #1
    Print #1, Date & " " & Time & " : " & ERR.Description & "(" & ERR.Number & ")"
    Close #1
    MsgBox ERR.Description
    Resume Next
'    Resume
    
End Sub

Private Sub Form_Resize()

If Me.Width < 8985 Then Exit Sub
If Me.Height < 7860 Then Exit Sub

'****************************************
frameContenedorx2.TOP = Me.ScaleHeight - 660
'***********************************************
frameContenedorx1.Left = Me.ScaleWidth - 2715
frameContenedorx1.TOP = Me.ScaleHeight - 4980
'***********************************************
Lista.Width = Me.ScaleWidth - 2865
'Lista.Height = Me.ScaleHeight - 3195
Lista.TOP = Me.ScaleHeight - 4965
'***********************************************
dgInput.Width = Me.ScaleWidth - 210
dgInput.Height = Me.ScaleHeight - 5025

End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSINPUT = Nothing
    Set RSCNPT = Nothing
    Set RSAUX = Nothing
End Sub

Private Sub IMAGE1_CLICK()
    'ELIMINACIÓN DEL TRABAJADOR
    On Error GoTo ERRMODULE
    If RSINPUT.EOF Then Exit Sub
    If MsgBox("Realmente desea quitar al trabajador o a los trabajadores seleccionados " & RSINPUT!NOMBRES, vbYesNo + vbQuestion) = vbYes Then
        CambiaPanelBD True
        Dim VARBMK As Variant
        For Each VARBMK In dgInput.SelBookmarks
            RSINPUT.Bookmark = VARBMK
            'SWDELETE SETEA PARA NO TENER ERRORES EN RSINPUT_MOVECOMPLETE
            SWDELETE = True
            DBSTARPLAN.Execute "DELETE FROM  [##ADELANTOS" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            DBSTARPLAN.Execute "DELETE FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            DBSTARPLAN.Execute "DELETE FROM  [##BOLSANT" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            DBSTARPLAN.Execute "DELETE FROM  [##PAGOSELIM" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            DBSTARPLAN.Execute "DELETE FROM  [##ADELELIM" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            RSINPUT.Delete
            SWDELETE = False
            RSINPUT.MovePrevious
            CALCULOTOTAL
            If RSINPUT.BOF Then
                If RSINPUT.RecordCount = 0 Then
                    MsgBox "ERROR, La lista no contiene trabajadores, se procederá a cerrar el formulario ", vbCritical
                    Unload Me
                Else
                    RSINPUT.MoveFirst
                End If
            End If
        Next
        xNumTrabs.Caption = " " & RSINPUT.RecordCount & " TRABAJADORES"
        CambiaPanelBD False
    End If
    Exit Sub
ERRMODULE:
    MsgBox "INCONSISTENCIA: Error al eliminar el registro " & ERR.Description
    Resume Next
End Sub
Private Sub CREARPLAN(Optional ByRef REG As Long)
    Dim TOTEGRE As Double, TOTINGR As Double
    Dim RSTRABPLAN As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim I As Integer, ORDEN As Integer
    Dim CONC As String
    If ExisteTablaAux(" [##TMPCREPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##TMPCREPLAN" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES VARCHAR(50),FECHAING DATETIME, CCOSTO VARCHAR(10), CARGO VARCHAR(30), BASICO  Numeric(20,2), FONDOPENS VARCHAR(2), CODCONCEP VARCHAR(15), DESCONCEP VARCHAR(40), ORDEN INT, MONTO  Numeric(20,2))"
    RSAUX.Open "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenKeyset, adLockReadOnly
    REG = RSAUX.RecordCount
    RSTRABPLAN.Open " [##TMPCREPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount = 0 Then
        MsgBox "No existe ningún registro para imprimir la Planilla"
        Exit Sub
    End If
    RSAUX.MoveFirst
    Dim RSCOLUM As New ADODB.Recordset
    RSCOLUM.Open "SELECT * FROM FORMARUBS,CONCEPTOS WHERE FORMARUBS.CONCEPTO=CONCEPTOS.CODIGO AND ID_FORMATO=" & VPTRASPRM & " ORDER BY CONCEPTOS.TIPO,CONCEPTOS.FILA", DBSYSTEM, adOpenKeyset
    Do While Not RSAUX.EOF
        RSCOLUM.Filter = "TIPO=0"
        If RSCOLUM.RecordCount > 0 Then
            RSCOLUM.MoveFirst
            Do While Not RSCOLUM.EOF
                ORDEN = ORDEN + 1
                For I = 0 To RSAUX.Fields.Count - 1
                    If Trim(RSAUX.Fields(I).Name) = RSCOLUM!Codigo And RSAUX.Fields(I).Value <> 0 Then
                        CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                             Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                        RSTRABPLAN.AddNew
                        Call LLENARRS(RSTRABPLAN, RSAUX)
                        RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                        RSTRABPLAN!DESCONCEP = CONC
                        RSTRABPLAN!ORDEN = ORDEN
                        RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                        RSTRABPLAN.Update
                        Exit For
                    End If
                Next
                RSCOLUM.MoveNext
            Loop
        End If
                
        RSCOLUM.Filter = "TIPO=1"
        If RSCOLUM.RecordCount > 0 Then
            RSCOLUM.MoveFirst
            Do While Not RSCOLUM.EOF
                ORDEN = ORDEN + 1
                For I = 0 To RSAUX.Fields.Count - 1
                    If Trim(RSAUX.Fields(I).Name) = RSCOLUM!Codigo And RSAUX.Fields(I).Value <> 0 Then
                        CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                             Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                        RSTRABPLAN.AddNew
                        Call LLENARRS(RSTRABPLAN, RSAUX)
                        RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                        RSTRABPLAN!DESCONCEP = CONC
                        RSTRABPLAN!ORDEN = ORDEN
                        RSTRABPLAN!MONTO = IIf(IsNull(RSAUX.Fields(I).Value), 0, RSAUX.Fields(I).Value)
                        RSTRABPLAN.Update
                        Exit For
                    End If
                Next
                RSCOLUM.MoveNext
            Loop
        End If
                    
        For I = 0 To RSAUX.Fields.Count - 1
            If UCase(Trim(RSAUX.Fields(I).Name)) = "OTROSING" Or UCase(Trim(RSAUX.Fields(I).Name)) = "TOTING" Then
                ORDEN = ORDEN + 1
                RSTRABPLAN.AddNew
                Call LLENARRS(RSTRABPLAN, RSAUX)
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                RSTRABPLAN!DESCONCEP = IIf(UCase(Trim(RSAUX.Fields(I).Name)) = "OTROSING", "OTROS INGRESOS", "TOTAL INGRESOS")
                RSTRABPLAN!ORDEN = ORDEN
                RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                TOTINGR = IIf(UCase(Trim(RSAUX.Fields(I).Name)) = "TOTING", IIf(IsNull(RSAUX.Fields(I).Value), 0, RSAUX.Fields(I).Value), 0)
                RSTRABPLAN.Update
                Exit For
            End If
        Next
                    
        RSCOLUM.Filter = "TIPO=2"
        If RSCOLUM.RecordCount > 0 Then
            RSCOLUM.MoveFirst
            Do While Not RSCOLUM.EOF
                ORDEN = ORDEN + 1
                For I = 0 To RSAUX.Fields.Count - 1
                    If Trim(RSAUX.Fields(I).Name) = RSCOLUM!Codigo And RSAUX.Fields(I).Value <> 0 Then
                        CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                             Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                        RSTRABPLAN.AddNew
                        Call LLENARRS(RSTRABPLAN, RSAUX)
                        RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                        RSTRABPLAN!DESCONCEP = CONC
                        RSTRABPLAN!ORDEN = ORDEN
                        RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                        RSTRABPLAN.Update
                        Exit For
                    End If
                Next
                RSCOLUM.MoveNext
            Loop
        End If
        
        For I = 0 To RSAUX.Fields.Count - 1
            If Trim(UCase(RSAUX.Fields(I).Name)) = "OTROSEGR" Or Trim(UCase(RSAUX.Fields(I).Name)) = "ADELANTOS" Or Trim(UCase(RSAUX.Fields(I).Name)) = "TOTEGR" Then
                ORDEN = ORDEN + 1
                RSTRABPLAN.AddNew
                Call LLENARRS(RSTRABPLAN, RSAUX)
                RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                Select Case Trim(UCase(RSAUX.Fields(I).Name))
                   Case "OTROSEGR": CONC = "OTROS EGRESOS"
                   Case "ADELANTOS": CONC = "ADELANTOS"
                   Case "TOTEGR":
                        CONC = "TOTAL EGRESOS":
                        TOTEGRE = IIf(IsNull(RSAUX.Fields(I).Value), 0, RSAUX.Fields(I).Value)
                End Select
                RSTRABPLAN!DESCONCEP = CONC
                RSTRABPLAN!ORDEN = ORDEN
                RSTRABPLAN!MONTO = IIf(IsNull(RSAUX.Fields(I).Value), 0, RSAUX.Fields(I).Value)
                RSTRABPLAN.Update
                Exit For
            End If
        Next
        
        RSCOLUM.Filter = "TIPO=3"
        If RSCOLUM.RecordCount > 0 Then
            RSCOLUM.MoveFirst
            Do While Not RSCOLUM.EOF
                ORDEN = ORDEN + 1
                For I = 0 To RSAUX.Fields.Count - 1
                    If Trim(RSAUX.Fields(I).Name) = RSCOLUM!Codigo And RSAUX.Fields(I).Value <> 0 Then
                        CONC = DevuelveValor("SELECT NOMBRE FROM CONCEPTOS WHERE CODIGO='" & _
                                             Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                        RSTRABPLAN.AddNew
                        Call LLENARRS(RSTRABPLAN, RSAUX)
                        RSTRABPLAN!CODCONCEP = Trim(RSAUX.Fields(I).Name)
                        RSTRABPLAN!DESCONCEP = CONC
                        RSTRABPLAN!ORDEN = ORDEN
                        RSTRABPLAN!MONTO = RSAUX.Fields(I).Value
                        RSTRABPLAN.Update
                        Exit For
                    End If
                Next
                RSCOLUM.MoveNext
            Loop
         End If
        'CALCULANDO EL NETO
        ORDEN = ORDEN + 1
        RSTRABPLAN.AddNew
        Call LLENARRS(RSTRABPLAN, RSAUX)
        RSTRABPLAN!CODCONCEP = "NETO"
        RSTRABPLAN!DESCONCEP = "NETO A PAGAR"
        RSTRABPLAN!ORDEN = ORDEN
        RSTRABPLAN!MONTO = TOTINGR - TOTEGRE
        RSTRABPLAN.Update
        ORDEN = 0
        RSAUX.MoveNext
    Loop
End Sub
Private Sub LLENARRS(RS As ADODB.Recordset, RS2 As ADODB.Recordset)
    RS!CODTRAB = RS2!CODTRAB
    RS!NOMBRES = RS2!NOMBRES
    RS!FECHAING = CDate(RS2!FECHAING)
    RS!CCosto = RS2!CODCCOSTO
    RS!CARGO = RS2!CARGO
    RS!BASICO = RS2!BASICO
    RS!FONDOPENS = RS2!CODAFP
End Sub

Private Sub IMAGE2_Click()
    CambiaPanelBD True
    Dim REG As Long
    Screen.MousePointer = 11
    Call CREARPLAN(REG)
    With Reporte
        .Reset
        '.LogOnServer "pdssql.dll", VGL_SERVERREP, "MARFICE_PP", "SOPORTE", "SOPORTE"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .WindowTitle = "PLAN0067.RPT - IMPRESION DE HOJA DE TRABAJO"
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0067.RPT"
        .StoredProcParam(0) = " [##TMPCREPLAN" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XCABEZA='PLANILLA : " & Me.Caption & "'"
        .Formulas(1) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        .Formulas(3) = "XREG='" & Str(REG) & "'"
        If .Status <> 2 Then .PrintReport
    End With
    Screen.MousePointer = 1
    CambiaPanelBD False
End Sub

Private Sub IMGBUSCAR_Click()
    XBUSCAR_Click
End Sub

Private Sub LABEL1_Click()
    IMAGE1_CLICK
End Sub

Private Sub Label2_Click()
    IMAGE2_Click
End Sub

Public Sub GRABARDATOS()
    Screen.MousePointer = 11
    If GetSetting(App.CompanyName, "PLANILLAS", "Nando", "NO") <> "HOLA" Then On Error GoTo ERRMODULE
    If Not CONSISTENCIA Then Exit Sub
    Dim RSCALC As New ADODB.Recordset
    Dim RSBOL As New ADODB.Recordset
    'PASAMOS A ELIMINAR LOS MOVIMIENTOS DE BOLETAS ANTERIORES Y QUE SE ENCUENTREN EL LA TABLA
    If ExisteTablaAux(" [##BOLSANT" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DELETE FROM " & REGINPUT.MOV_TABLE & " WHERE CODNOMBOL=" & REGINPUT.Codigo & " AND INUMBOL IN (SELECT INUMBOL FROM  [##BOLSANT" & VGL_COMPUTER & "] )"
    'ABRIMOS LA TABLA DE CALCULOS DE PLANILLA CALCINPUT
    RSCALC.Open "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    'BORRAMOS LAS BOLETAS YA PROCESADAS
    DBSYSTEM.Execute "DELETE FROM " & REGINPUT.BOL_TABLE & " WHERE CODNOMBOL=" & REGINPUT.Codigo & " AND CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )"
    'APLICAMOS SQL PARA ESCRIBIR MAS RÁPIDO Y SEGURO
    Do While Not RSCALC.EOF
        DBSYSTEM.Execute "INSERT INTO " & REGINPUT.BOL_TABLE & " (NUMBOL,CODNOMBOL,CODTRAB,FECHA,CODAFP,CCOSTO,TIPOPLAN,BASICO,SUMAAFP,SUMASALUD,SUMAIES,SUMARENTA,SUMASCTR,SUMACTS,SUMAGRAT,SUMAVAC,TOTING,TOTEGR,HORASTRAB,HORASEXTRAS,RENTA5TA,XREDONDEO) VALUES (1," & REGINPUT.Codigo & ",'" & RSCALC!CODTRAB & "'," & DateSQL(Date) & ",'" & RSCALC!CODAFP & "','" & RSCALC!CODCCOSTO & "',1," & RSCALC!BASICO & "," & IIf(IsNull(RSCALC!SUMAAFP), "NULL", RSCALC!SUMAAFP) & "," & IIf(IsNull(RSCALC!SUMASALUD), "NULL", RSCALC!SUMASALUD) & "," & IIf(IsNull(RSCALC!SUMAIES), "NULL", RSCALC!SUMAIES) & "," & IIf(IsNull(RSCALC!SUMARENTA), "NULL", RSCALC!SUMARENTA) & "," & _
        IIf(IsNull(RSCALC!SUMASCTR), "NULL", RSCALC!SUMASCTR) & "," & IIf(IsNull(RSCALC!SUMACTS), "NULL", RSCALC!SUMACTS) & "," & IIf(IsNull(RSCALC!SUMAGRAT), "NULL", RSCALC!SUMAGRAT) & "," & IIf(IsNull(RSCALC!SUMAVAC), "NULL", RSCALC!SUMAVAC) & "," & IIf(IsNull(RSCALC!TOTING), "NULL", RSCALC!TOTING) & "," & IIf(IsNull(RSCALC!TOTEGR), "NULL", RSCALC!TOTEGR) & "," & RSCALC.Fields("_HORAST") & "," & RSCALC.Fields("_HOREXTRAS") & "," & RSCALC.Fields("_QUINTACAT") & "," & RSCALC.Fields("XREDONDEO") & ")"
        'CAPTURA DEL CÓDIGO DE LA ÚLTIMA BOLETA AGREGADA
        Set RSAUX = Nothing
        RSAUX.Open "SELECT INUMBOL FROM " & REGINPUT.BOL_TABLE & " WHERE CODNOMBOL=" & REGINPUT.Codigo & " AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM, adOpenStatic
        'AGREGAMOS LOS RUBROS MAYORES DE 0 A LA TABLA DE MOVIMIENTOS
        '--------------------------------------------------------------------------
        RSCNPT.MoveFirst
        Do While Not RSCNPT.EOF
            If Not DevuelveValor("SELECT PERMITE FROM CONCEPTOS WHERE CODIGO='" & Trim$(RSCNPT!Codigo) & "'", DBSYSTEM) Then
                If RSCALC.Fields(Trim$(RSCNPT!Codigo)).Value > 0 Then
                    DBSYSTEM.Execute "INSERT INTO " & REGINPUT.MOV_TABLE & " (INUMBOL, CONCEPTO, MONTO, CODNOMBOL) VALUES (" & RSAUX!INUMBOL & ",'" & RSCNPT!Codigo & "'," & RSCALC.Fields(Trim$(RSCNPT!Codigo)).Value & "," & REGINPUT.Codigo & ")"
                End If
              Else: DBSYSTEM.Execute "INSERT INTO " & REGINPUT.MOV_TABLE & " (INUMBOL, CONCEPTO, MONTO, CODNOMBOL) VALUES (" & RSAUX!INUMBOL & ",'" & RSCNPT!Codigo & "'," & ESNULO(RSCALC.Fields(Trim$(RSCNPT!Codigo)).Value, 0) & "," & REGINPUT.Codigo & ")"
            End If
            RSCNPT.MoveNext
        Loop
        RSCALC.MoveNext
    Loop
    'GRABAR ADELANTOS DE REMUNERACIONES
    If CalcPlan.Check1.Value = 1 Then
        'ASIGNAMOS ADELANTOS A LA BOLETA
        DBSYSTEM.Execute "UPDATE ADEL" & REGSISTEMA.ANNO & " SET NOMBOL=" & REGINPUT.Codigo & " WHERE CODIGO IN (SELECT CODIGO FROM  [##ADELANTOS" & VGL_COMPUTER & "] )"
        'BORRAMOS LAS ASIGNACIONES QUITADAS POR EL USUARIO
        DBSYSTEM.Execute "UPDATE ADEL" & REGSISTEMA.ANNO & "  SET NOMBOL=0 WHERE CODIGO IN (SELECT CODIGO FROM  [##ADELELIM" & VGL_COMPUTER & "] )"
    End If
    If CalcPlan.Check3.Value = 1 Then
        Set RSAUX = Nothing
        RSAUX.Open " [##PAGOSELIM" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenStatic
        'DEVOLVEMOS LO DEBITADO PARA AQUELLOS PAGOS ELIMINADOS
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO+" & RSAUX!CUOTA & " WHERE CODMOV=" & RSAUX!CODMOV
            'ELIMINAMOS LOS PAGOS QUE HAN SIDO CANCELADOS
            DBSYSTEM.Execute "DELETE FROM PAGOSCTA WHERE CODMOV=" & RSAUX!CODMOV & " AND CODNOMBOL=" & REGINPUT.Codigo
            RSAUX.MoveNext
        Loop
        'GRABAMOS FINALMENTE LAS CUENTAS CORRIENTES
        Set RSAUX = Nothing
        Set RSAUX = Nothing
        Dim Diferencia As Integer
        If ExisteTablaAux(" [##PAGOSCTACTE" & VGL_COMPUTER & "] ") Then
                Dim RsAdelAux As ADODB.Recordset
        RSAUX.Open "SELECT * FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODNOMBOL=0", DBSTARPLAN, adOpenStatic
        Do While Not RSAUX.EOF
            Set RsAdelAux = New ADODB.Recordset
            RsAdelAux.Open "SELECT * FROM PAGOSCTA WHERE CODMOV=" & RSAUX!CODMOV & " AND CODTRAB='" & RSAUX!CODTRAB & "' AND TIPO=" & RSAUX!TIPOGRUPO & " AND CODNOMBOL=" & REGINPUT.Codigo, DBSYSTEM, adOpenKeyset, adLockOptimistic
            Diferencia = 0
            If RsAdelAux.RecordCount > 0 Then
                Diferencia = RsAdelAux!MONTO
            End If
            'DEBITAMOS AL SALDO
            DBSYSTEM.Execute "UPDATE MOVICTA SET SALDO=SALDO-" & (RSAUX!CUOTA - Diferencia) & " WHERE CODMOV=" & RSAUX!CODMOV
            DBSYSTEM.Execute "INSERT INTO PAGOSCTA (CODMOV, NUMBOL, CODNOMBOL, TIPOBOLETA, MONTO, DOLAR, CODTRAB, TIPO,SECUENCIA) VALUES (" & RSAUX!CODMOV & ",1," & REGINPUT.Codigo & ",'B'," & RSAUX!CUOTA & ",0,'" & RSAUX!CODTRAB & "'," & RSAUX!TIPOGRUPO & "," & RSAUX!ULTSECU & ")"
            RSAUX.MoveNext
        Loop
        End If
    End If
    'SE GRABAN LAS VACACIONES
    '------------------------
    Set RSAUX = Nothing
    RSAUX.Open "SELECT CODIGO FROM HISTOVAC WHERE CODTRAB IN (SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] )", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount > 0 Then
        Do While Not RSAUX.EOF
            DBSYSTEM.Execute "UPDATE HISTOVAC SET CERRADO=1 WHERE CODIGO=" & RSAUX!Codigo
            RSAUX.MoveNext
        Loop
    End If
    Set RSCALC = Nothing
    Set RSBOL = Nothing
    'APLICACION DE REDONDEO
    If REGINPUT.REDONDEO Then
        DBSYSTEM.Execute "UPDATE TRABAJADORES SET TRABAJADORES.XREDONDEO=B.XREDONDEO FROM TRABAJADORES A, [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  B WHERE A.CODTRAB=B.CODTRAB"
    End If
    'ACTUALIZA EL ACUMULADO DE LA QUINTA CATEGORIA
            Dim RS_AUX_QU As New ADODB.Recordset
            Set RS_AUX_QU = New ADODB.Recordset
            RS_AUX_QU.Open "SELECT * FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
            If RS_AUX_QU.RecordCount > 0 Then
                While Not RS_AUX_QU.EOF 'BORRA LOS YA EXISTENTES EN ESA MISMA FECHA
                    DBSYSTEM.Execute "DELETE FROM HIST5TA WHERE MES=" & Month(REGINPUT.MESACTIVO) & " AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RS_AUX_QU!CODTRAB & "'", X
                    RS_AUX_QU.MoveNext
                Wend
            End If
            'INSERTA LOS NUEVOS REGSITROS
            DBSYSTEM.Execute "INSERT INTO HIST5TA SELECT * FROM  [##TMPQUINTA" & VGL_COMPUTER & "]  ", X
        Screen.MousePointer = 1
        MsgBox "LOS DATOS SE GRABARON SATISFACTORIAMENTE. SE PROCEDERÁ A ABANDONAR LA PRESENTE VENTANA", vbInformation
    Unload Me
    Exit Sub
ERRMODULE:
    MsgBox "INCONSISTENCIA: " & ERR.Description
    Resume Next
    Resume
End Sub

Private Sub Lista_DblClick()
    On Error GoTo ERRNOLIST
    Set XITEM = Lista.SelectedItem
    If Left(XITEM.SubItems(1), 3) = "(*)" Then
        If MsgBox("Desea quitar: " & XITEM.SubItems(1), vbQuestion + vbYesNo) = vbNo Then Exit Sub
        If XITEM.Tag = 2 Then
            DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET ADELANTO=ADELANTO-" & XITEM.SubItems(2) & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            MsgBox "Se ha eliminado temporalmente el Adelanto de Remuneraciones", vbInformation
            DBSTARPLAN.Execute "INSERT INTO  [##ADELELIM" & VGL_COMPUTER & "]  VALUES('" & RSINPUT!CODTRAB & "'," & XITEM.Text & ")"
            DBSTARPLAN.Execute "DELETE FROM  [##ADELANTOS" & VGL_COMPUTER & "]  WHERE CODIGO=" & XITEM.Text
        Else
            If Left(XITEM.Text, 1) = "I" Then 'SI ES INGRESO
                DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSING=OTROSING-" & XITEM.SubItems(2) & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            Else
                DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET OTROSEGR=OTROSEGR-" & XITEM.SubItems(2) & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            End If
            MsgBox "Se ha eliminado temporalmente el Pago de Cuenta Corriente. Los cambios se aceptarán una vez que se halla grabado ", vbInformation
            DBSTARPLAN.Execute "INSERT INTO  [##PAGOSELIM" & VGL_COMPUTER & "]  VALUES ('" & RSINPUT!CODTRAB & "'," & Right(XITEM.Text, Len(XITEM.Text) - 2) & "," & XITEM.SubItems(2) & ")"
            DBSTARPLAN.Execute "DELETE FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODMOV=" & Right(XITEM.Text, Len(XITEM.Text) - 2)
        End If
        CALCULOTOTAL
        RSINPUT.MOVE 0
        Exit Sub
    End If
    Select Case XITEM.Tag
        Case 1
            MsgBox "Fórmula: " & DevuelveValor("SELECT FORMULA FROM CONCEPTOS WHERE CODIGO='" & XITEM.Text & "'", DBSYSTEM) & Chr(13) & Chr(10) & "COMENTARIO: " & DevuelveValor("SELECT COMENTARIO FROM CONCEPTOS WHERE CODIGO='" & XITEM.Text & "'", DBSYSTEM), vbInformation, "FORMULA DE: " & XITEM.SubItems(1)
        Case 2 'SI ES UN ADELANTO, PREGUNTAR SI DESEA QUITARLO
            frCmbInp.Show 1
        Case 3
            frCmbInp.Show 1
    End Select
    Exit Sub
ERRNOLIST:
    Exit Sub
End Sub

Private Sub RSINPUT_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If SWDELETE Then Exit Sub
    If PRECORDSET.EOF Or PRECORDSET.BOF Then Exit Sub
    Call CALCULOXREGISTRO(PRECORDSET)
End Sub
Private Sub CALCULOXREGISTRO(PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRMODULE
    'CAMBIAPANELBD TRUE
    xCodTrab.Tag = "" & RSINPUT!CODTRAB
    Dim X As Integer
    For X = Lista.ListItems.Count To 1 Step -1
        Lista.ListItems(X).SubItems(2) = "0.00 "
        If Lista.ListItems(X).Tag <> 1 Then Lista.ListItems.Remove X
    Next
    Set RSAUX = Nothing
    RSAUX.Open "SELECT * FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE CODTRAB='" & PRECORDSET!CODTRAB & "'", DBSTARPLAN, adOpenStatic
    If Not RSAUX.EOF Then
        For X = Lista.ListItems.Count To 1 Step -1
            Lista.ListItems(X).SubItems(2) = Format(RSAUX.Fields(Lista.ListItems(X).Text).Value, "0.00 ")
        Next
    End If
    xBasico.Caption = Format(RSAUX!BASICO, "0.00 ")
    xTotIng.Caption = Format(RSAUX!TOTING, "0.00 ")
    xTotEgr.Caption = Format(RSAUX!TOTEGR, "0.00 ")
    XREDONDEO.Caption = Format(IIf(IsNull(RSAUX!XREDONDEO), 0, RSAUX!XREDONDEO), "0.00 ")
    xArea.Caption = RSAUX!CODAFP & " "
    xCCosto.Caption = RSAUX!CODCCOSTO & " "
    xNeto.Caption = Format(RSAUX!TOTING - RSAUX!TOTEGR + IIf(REGINPUT.REDONDEO, RSAUX!XREDONDEO, 0), "0.00 ")
    If (RSAUX!TOTING - RSAUX!TOTEGR) < 0 Then xNeto.BackColor = &HFF& Else xNeto.BackColor = &H80000009
    Set RSAUX = Nothing
    If CalcPlan.Check1.Value = 1 Then 'SI CARGAR ADELANTOS
        'REVISAR
        Set RSAUX = Nothing
        RSAUX.Open "SELECT MES,MONTO,CODIGO,NOMBOL FROM  [##ADELANTOS" & VGL_COMPUTER & "]  WHERE CODTRAB='" & PRECORDSET!CODTRAB & "'", DBSTARPLAN, adOpenStatic
        Do While Not RSAUX.EOF
            Set XITEM = Lista.ListItems.Add(, , ESNULO(RSAUX!Codigo, " "), , 5)
            XITEM.SubItems(1) = IIf(RSAUX!NOMBOL <> 0, "(*) ", "") & "Adelanto de Rem. " & AMESES(Month(RSAUX!MES))
            XITEM.SubItems(2) = Format(RSAUX!MONTO, "0.00")
            XITEM.Tag = 2
            RSAUX.MoveNext
        Loop
    End If
    If CalcPlan.Check4.Value = 1 Then 'SI CARGAR CUENTAS CORRIENTES
        Set RSAUX = Nothing
        RSAUX.Open "SELECT DESCRIPCION,TIPOGRUPO,CODMOV,CUOTA,CODNOMBOL FROM  [##PAGOSCTACTE" & VGL_COMPUTER & "]  WHERE CODTRAB='" & PRECORDSET!CODTRAB & "'", DBSTARPLAN, adOpenStatic
        If ESNULO(RSAUX!CUOTA, 0) <> 0 Then
            Do While Not RSAUX.EOF
                Set XITEM = Lista.ListItems.Add(, , IIf(RSAUX!TIPOGRUPO = 1, "I-" & RSAUX!CODMOV, "E-" & RSAUX!CODMOV), , IIf(RSAUX!TIPOGRUPO = 1, 6, 7))
                XITEM.SubItems(1) = IIf(RSAUX!CODNOMBOL = 0, "", "(*) ") & RSAUX!DESCRIPCION
                XITEM.SubItems(2) = Format(RSAUX!CUOTA, "0.00")
                XITEM.Tag = 3
                RSAUX.MoveNext
            Loop
        End If
    End If
    Lista.Refresh
    'CAMBIAPANELBD FALSE
    Exit Sub
ERRMODULE:
    CambiaPanelBD False
    MsgBox "INCONSISTENCIA: " & ERR.Description
    Resume Next
'    Resume
    CambiaPanelBD True
End Sub
Public Sub CALCULOTOTAL(Optional FILTRAR As String = "")
    'CALCULO DE FORMULAS TIPO INGRESO
    If GetSetting(App.CompanyName, "PLANILLAS", "Nando", "NO") <> "HOLA" Then On Error GoTo ERRMODULE
    With RSCNPT
        .MoveFirst
        If ESNULO(InStr(!FORMULA, "@"), 0) = 0 Then
            Do While Not .EOF
                If !TIPO <= 1 And Not !ESESCRITO And Not IsNull(!FORMULA) Then 'SI TIPO ES INGRESO
                    DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=" & !FORMULA & IIf(Trim(!CRITERIO) = "", "", " WHERE " & !CRITERIO)
                    If !CRITERIO <> "" Then DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=0 " & IIf(Trim(!CRITERIO) = "", "", " WHERE NOT (" & !CRITERIO & ")")
                End If
                .MoveNext
            Loop
        End If
    End With
    '--------------------------------------
    'CALCULO DE LAS SUMAS DE AFECTOS A
    '--------------------------------------
    
    
    DBSTARPLAN.Execute STRCALCSUM
    
    '--------------------------------------
    'Calculo del Impuesto de 5ta. Categoria
    '--------------------------------------
    
Dim TOTALPERCIBIDO
Dim PROYECTADOFINAL
Dim TOTALRENTAPERCIBIR
Dim RENTAAFECTA
Dim IMPUESTOANUAL
Dim ACUMULADO
Dim SALDO
Dim MONTORETENER
Dim RETENCIONANTERIOR
Dim SQL As String
Dim UIT1 As Double, UIT2 As Double, VALUIT As Double
Dim UIT3 As Double, UIT4 As Double, Porc1 As Double, Porc2 As Double, Porc3 As Double
    
    If ExisteTablaAux(" [##TMPQUINTA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPQUINTA" & VGL_COMPUTER & "] "
     DBSYSTEM.Execute "CREATE TABLE  [##TMPQUINTA" & VGL_COMPUTER & "]  (MES INT, ANNO VARCHAR(4), CODTRAB VARCHAR(8), NOMBRES VARCHAR(100), [TOTAL PERCIBIDO]  Numeric(20,2), [PROYECTADO FIN AÑO]  Numeric(20,2), [TOTAL RENTA PERCIBIR]  Numeric(20,2), [RENTA AFECTA]  Numeric(20,2), [IMPUESTO ANUAL]  Numeric(20,2), [ACUMULADO]  Numeric(20,2), SALDO  Numeric(20,2), [MONTO RETENER]  Numeric(20,2), [RENTENCION ANTERIOR]  Numeric(20,2))"
    
    
    Dim XNUMM As Byte, VINGMIN As Double, M1 As Double, M2 As Double, S1 As Double, S2 As Double, M3 As Double, M4 As Double
    Dim RSAUX1 As New ADODB.Recordset, RSCALC As New ADODB.Recordset, RENTADEDUC As Double, XVALMES As Integer
    Dim RSTRAB As New ADODB.Recordset, XVALOPCION As Double, XC As Double
    Dim MESACUMULADO As Integer
    Dim MINIMO As Double
    XNUMM = Month(REGINPUT.MESACTIVO) ' NUMERO DEL MES
    'AHORA USAMOS RSAUX1 PARA ABRIR EL RECORDSET DE IMPUESTOS ANTERIORES
    Set RSAUX1 = New ADODB.Recordset
    RSAUX1.Open "CONFIG5TA", DBSYSTEM, adOpenStatic
    VALUIT = RSAUX1!VALORUIT
    UIT1 = RSAUX1!UIT1
    UIT2 = RSAUX1!UIT2
    UIT3 = RSAUX1!UIT3
    UIT4 = RSAUX1!UIT4
    Porc1 = RSAUX1!PORCENTAJE
    Porc2 = RSAUX1!PORCENTAJE2
    Porc3 = RSAUX1!PORCENTAJE3
    
    
    Set RSCALC = New ADODB.Recordset
    RSCALC.Open "SELECT * FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE AFECTOQUINTA=1", DBSYSTEM, adOpenKeyset, adLockOptimistic 'SOLO AQUELLOS QUE ESTEN AFECTOS A QUINTA
    'RSCALC.Open "SELECT * FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic 'SOLO AQUELLOS QUE ESTEN AFECTOS A QUINTA
    RENTADEDUC = RSAUX1!NUMUIT * RSAUX1!VALORUIT 'AHORA SON 7 UIT * 3000
    XVALMES = Val(RSAUX1("MES" & Format(XNUMM, "00"))) 'NUMERO QUE DIVIDIRA
    MINIMO = RSAUX1!VALORMIN
    
    Select Case XNUMM 'SIRVE PARA SACAR EL ACUMULADO DE MESES ANTERIORES
        Case 1
            MESACUMULADO = RSAUX1!ACUMULA01
        Case 2
            MESACUMULADO = RSAUX1!ACUMULA02
        Case 3
            MESACUMULADO = RSAUX1!ACUMULA03
        Case 4
            MESACUMULADO = RSAUX1!ACUMULA04
        Case 5
            MESACUMULADO = RSAUX1!ACUMULA05
        Case 6
            MESACUMULADO = RSAUX1!ACUMULA06
        Case 7
            MESACUMULADO = RSAUX1!ACUMULA07
        Case 8
            MESACUMULADO = RSAUX1!ACUMULA08
        Case 9
            MESACUMULADO = RSAUX1!ACUMULA09
        Case 10
            MESACUMULADO = RSAUX1!ACUMULA10
        Case 11
            MESACUMULADO = RSAUX1!ACUMULA11
        Case 12
            MESACUMULADO = RSAUX1!ACUMULA12
    End Select
    Dim XMESVAR As Integer
    
    XVALOPCION = 12 - XNUMM + 1 'MES PROYECTADA
    XMESVAR = 2
    If XNUMM > 7 And XNUMM <= 12 Then XMESVAR = 1
    
    Dim L As Double
    Dim M_ACUM_ANT_DELMES As Double
    Dim MONTOVAR As Double, T_QUINTA_ANT As Double
    
    
    If FILTRAR <> "" Then RSCALC.Filter = "CODTRAB='" & FILTRAR & "'"
    Do While Not RSCALC.EOF
        Set RSTRAB = New ADODB.Recordset
        'RECUPERA UN INGRESO EXTRA DEL TRABAJADOR
        RSTRAB.Open "SELECT TOTALEXTRA FROM TRABAJADORES WHERE CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
        If RSTRAB.RecordCount > 0 Then
            M_ACUM_ANT_DELMES = DevuelveValor("SELECT SUM([TOTAL PERCIBIDO]) FROM HIST5TA WHERE MES<=" & Str(XNUMM - 1) & "  AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)
            MONTOVAR = GET_HISTVAR(SUMAVAR, RSCALC!CODTRAB, CDate(REGINPUT.MESACTIVO))
            
            'M1 es la Proyeccion del Monto
            'L es la Renta Basica
            M1 = ((RSCALC!SUMARENTA + RSTRAB.Fields(0) + MONTOVAR) * XVALOPCION) + M_ACUM_ANT_DELMES + ((RSCALC!SUMARENTA + RSTRAB.Fields(0)) * XMESVAR)
            L = (M1 + RSCALC!T4) - RENTADEDUC
            'Calculando el 1er. Tope
            If M1 > RENTADEDUC And M1 <= (VALUIT * UIT1) Then
                PROYECTADOFINAL = M1
                T_QUINTA_ANT = DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE MES>=" & (XNUMM - MESACUMULADO) & " AND MES<=" & (XNUMM - 1) & "  AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)
                TOTALPERCIBIDO = RSCALC!SUMARENTA + MONTOVAR
                RENTAAFECTA = (TOTALRENTAPERCIBIR) - RENTADEDUC
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4)
                
                M2 = ((L * (Porc1 / 100)) - T_QUINTA_ANT) / XVALMES
                IMPUESTOANUAL = (RENTAAFECTA) * (Porc1 / 100)
                ACUMULADO = T_QUINTA_ANT
                SALDO = IMPUESTOANUAL - ACUMULADO
                MONTORETENER = M2
                RETENCIONANTERIOR = IIf(MESACUMULADO = 0, 0, DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM))
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=" & IIf((M2 < 0), (0), (M2)) & " WHERE CODTRAB='" & RSCALC!CODTRAB & "'"
            ElseIf M1 > RENTADEDUC And M1 > (VALUIT * UIT1) Then
                PROYECTADOFINAL = M1
                T_QUINTA_ANT = DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE MES>=" & (XNUMM - MESACUMULADO) & " AND MES<=" & (XNUMM - 1) & "  AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)
                TOTALPERCIBIDO = RSCALC!SUMARENTA + MONTOVAR
                RENTAAFECTA = (TOTALRENTAPERCIBIR) - RENTADEDUC
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4)
            'Calculando el 2do. Tope
                M3 = (UIT2 * VALUIT) * (Porc1 / 100)
                If L - (UIT2 * VALUIT) > (VALUIT * UIT1) And L - (UIT2 * VALUIT) <= (UIT3 * VALUIT) Then
                    M4 = (L - (UIT2 * VALUIT)) * (Porc2 / 100)
                ElseIf L - (UIT2 * VALUIT) > (UIT4 * VALUIT) Then
                    M4 = (L - (UIT2 * VALUIT)) * (Porc3 / 100)
                Else
                    M4 = (L - (UIT2 * VALUIT)) * (Porc1 / 100)
                End If
                M2 = ((M3 + M4) - T_QUINTA_ANT) / XVALMES
                IMPUESTOANUAL = (M3 + M4)
                ACUMULADO = T_QUINTA_ANT
                SALDO = IMPUESTOANUAL - ACUMULADO
                MONTORETENER = M2
                RETENCIONANTERIOR = IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=" & IIf((M2 < 0), (0), (M2)) & " WHERE CODTRAB='" & RSCALC!CODTRAB & "'"
            Else
                TOTALPERCIBIDO = RSCALC!SUMARENTA + MONTOVAR
                PROYECTADOFINAL = M1
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4)
                RENTAAFECTA = 0
                IMPUESTOANUAL = 0
                ACUMULADO = 0
                SALDO = 0
                MONTORETENER = 0
                RETENCIONANTERIOR = 0
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=0" & " WHERE CODTRAB='" & RSCALC!CODTRAB & "'"
            End If
            
            DBSYSTEM.Execute "INSERT INTO  [##TMPQUINTA" & VGL_COMPUTER & "]  (MES , ANNO , CODTRAB , NOMBRES , [TOTAL PERCIBIDO] , [PROYECTADO FIN AÑO] , [TOTAL RENTA PERCIBIR], [RENTA AFECTA] , [IMPUESTO ANUAL] , [ACUMULADO] , SALDO , [MONTO RETENER] , [RENTENCION ANTERIOR] ) VALUES (" & _
                "" & Month(REGINPUT.MESACTIVO) & ",'" & Year(REGINPUT.MESACTIVO) & "','" & RSCALC!CODTRAB & "','" & RSCALC!NOMBRES & "'," & TOTALPERCIBIDO & "," & PROYECTADOFINAL & _
                "," & TOTALRENTAPERCIBIR & "," & RENTAAFECTA & "," & IMPUESTOANUAL & "," & ACUMULADO & "," & SALDO & "," & _
                MONTORETENER & "," & RETENCIONANTERIOR & ")"
                
        End If
        Set RSTRAB = Nothing
        RSCALC.MoveNext
    Loop
    Set RSAUX1 = Nothing
    Set RSAUX = Nothing
    Set RSCALC = Nothing
    Dim RSNTRAB As New ADODB.Recordset
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            'CALCULO DE EGRESOS Y APORTACIONES, SI ES QUE TIENEN FORMULA. GENERALMENTE LAS APORTACIONES NO LO TIENEN PERO SE INCLUYEN
            If !TIPO > 1 And Not !ESESCRITO And Not IsNull(!FORMULA) Then 'SI TIPO ES INGRESO
                If ESNULO(InStr(!FORMULA, "@"), 0) = 0 Then
                    If Left(!FORMULA, 1) = "_" Then
                        DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=[" & !FORMULA & "]" & IIf(Len(Trim(!CRITERIO)) = 0, "", " WHERE " & !CRITERIO)
                        If Len(Trim(!CRITERIO)) > 0 Then DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=0 " & IIf(!CRITERIO = "", "", " WHERE NOT (" & !CRITERIO & ")")
                    Else
                        DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=" & !FORMULA & IIf(Len(Trim(!CRITERIO)) = 0, "", " WHERE " & !CRITERIO)
                        If Len(Trim(!CRITERIO)) > 0 Then DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=0 " & IIf(Len(Trim(!CRITERIO)) = 0, "", " WHERE NOT (" & !CRITERIO & ")")
                    End If
                  Else
                    Set RSNTRAB = New ADODB.Recordset
                    RSNTRAB.Open "SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]", DBAUXCOM, adOpenKeyset, adLockReadOnly
                    If FILTRAR <> "" Then RSNTRAB.Filter = "CODTRAB='" & FILTRAR & "'"
                    Do While Not RSNTRAB.EOF
                        DBAUXCOM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=" & RPLCADENA(!FORMULA, RSNTRAB!CODTRAB, REGINPUT.FECHAFIN, REGINPUT.Codigo) & " WHERE CODTRAB='" & RSNTRAB!CODTRAB & "'"
                        RSNTRAB.MoveNext
                    Loop
                End If
            End If
            .MoveNext
        Loop
    End With
    'CALCULO DEL TOTAL EGRESOS
    '-------------------------
    'YA SE INCLUYEN EN ESTE PROCESO LOS OTROS EGRESOS
    Dim CADSUMAS As String
    CADSUMAS = "0+OTROSEGR+ADELANTO"
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO = 2 Then
                'PARA EL TOTAL DE INGRESOS
                CADSUMAS = CADSUMAS & "+" & !Codigo
            End If
            .MoveNext
        Loop
        DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET TOTEGR=" & CADSUMAS
    End With
    If Not ExisteCampo("XREDONDEO", REGINPUT.BOL_TABLE, DBSYSTEM) Then
        'COMPATIBILIDAD CON VERSIONES ANTERIORES
        DBSYSTEM.Execute "ALTER TABLE " & REGINPUT.BOL_TABLE & " ADD XREDONDEO  Numeric(20,2)"
        DBSYSTEM.Execute "UPDATE " & REGINPUT.BOL_TABLE & " SET XREDONDEO=0"
    End If
    'APLICA REDONDEO
    If REGINPUT.REDONDEO Then
        DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET XREDONDEO=CASE WHEN ([TOTING]-[TOTEGR])-ROUND([TOTING]-[TOTEGR],0,1)>=0.5 THEN " _
            & "1-(([TOTING]-[TOTEGR])-(ROUND([TOTING]-[TOTEGR],0,1))) ELSE -1*(([TOTING]-[TOTEGR])-(ROUND([TOTING]-[TOTEGR],0,1))) END"
    End If
    CALCULOTOTALGEN
    Exit Sub
ERRMODULE:
    If RSCNPT.EOF Then
        MsgBox "INCONSISTENCIA (" & ERR.Number & "). Posible fórmula invalida en el Sistema: " & ERR.Description
    Else
        MsgBox "Error de Usuario en el Concepto de Remuneración: (" & RSCNPT!Codigo & ") " & RSCNPT!NOMBRE & "FÓRMULA: " & RSCNPT!FORMULA & Chr(13) & Chr(10) & " DETALLE DE ERROR: " & ERR.Description
    End If
    Resume Next
    Resume
End Sub

Private Sub RSINPUT_RECORDCHANGECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal CRECORDS As Long, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    On Error GoTo ERRMODULE
    'Exit Sub
    If SWDELETE Then Exit Sub
    If ADREASON = adRsnUpdate Then
        Call CALCULOTOTAL(PRECORDSET!CODTRAB)
    End If
    Exit Sub
ERRMODULE:
    MsgBox "INCONSISTENCIA: Posible registro invalido en windows: " & ERR.Description
    Resume Next
End Sub

Public Sub CALCULOTOTALGEN()
    On Error GoTo ERRMODULE
    Dim XSTR As String
    For X = Lista.ListItems.Count To 1 Step -1
        XSTR = Lista.ListItems(X).Text
        Set RSAUX = Nothing
        If Lista.ListItems(X).Tag = 1 Then
            RSAUX.Open "SELECT SUM(" & XSTR & ") AS TOTAL FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSTARPLAN, adOpenStatic
            Lista.ListItems(X).SubItems(3) = Format(RSAUX!TOTAL, "0.00 ")
        End If
    Next
    Set RSAUX = Nothing
    DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET XREDONDEO=0 WHERE XREDONDEO IS NULL"
    RSAUX.Open "SELECT SUM(TOTING) AS TOTAL1, SUM(TOTEGR-XREDONDEO) AS TOTAL2 FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSTARPLAN, adOpenStatic
    xTGIng.Caption = Format(RSAUX!Total1, "0.00 ")
    xTGEgr.Caption = Format(RSAUX!Total2, "0.00 ")
    xTGNeto.Caption = Format(RSAUX!Total1 - RSAUX!Total2, "0.00 ")
    Exit Sub
ERRMODULE:
    MsgBox "INCONSISTENCIA: POSIBLE REGISTRO INVÁLIDO EN WINDOWS: " & ERR.Description
    Resume Next
End Sub

Private Sub XBUSCAR_Click()
    CambiaPanelBD True
    Set RSAUX = RSINPUT.Clone
    frmComun.CONECTAR RSAUX
    CambiaPanelBD False
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        RSINPUT.MoveFirst
        RSINPUT.FIND "CODTRAB='" & VGUTIL(1) & "'"
    End If
End Sub

Public Sub REFRESCARTRAB()
    RSINPUT.MOVE 0
End Sub

Public Function CONSISTENCIA() As Boolean
    Set RSAUX = Nothing
    'CAPTURA DE LOS RESULTADOS NEGATIVOS
    RSAUX.Open "SELECT CODTRAB FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE TOTEGR>TOTING", DBSTARPLAN, adOpenStatic
    If RSAUX.RecordCount > 0 Then
        SWDELETE = True
        MsgBox "EXISTE PERSONAL CON TOTALES NETO A RECIBIR MENORES A CERO. EL RESULTADO DE LA PLANILLA NO PUEDE SER NEGATIVO", vbCritical
        RSINPUT.MoveFirst
        RSINPUT.FIND "CODTRAB='" & RSAUX!CODTRAB & "'"
        SWDELETE = False
        RSINPUT.MOVE 0
        CONSISTENCIA = False
        Screen.MousePointer = 1
    Else
        CONSISTENCIA = True
    End If
End Function

Private Function GET_HISTVAR(CADENA As String, CODTRAB As String, FECHA As Date) As Double
    If Len(CADENA) <= 11 Then Exit Function
    Dim I As Integer, j As Integer, CONT As Integer, NUMVECES As Integer
    Dim MESANO As String
    Dim RsTmp As New ADODB.Recordset
    Dim STRCADENA As String, CAD As String, CODCONCEPTO As String
    Dim TOTAL As Double, TOTALGEN As Double
    Dim AUXCAD As String
    Dim MONTO As Double
    Dim MOTOGEN As Double
    CAD = Right(CADENA, Len(CADENA) - 20)
    AUXCAD = CAD
    MOTOGEN = 0
    'RECORRIENDO LOS CONCEPTOS QUE SON VARIABLES Y ESTAN AFECTOS A QUINTA
    For j = 1 To BusCad("+", CAD)
        TOTAL = 0
        CODCONCEPTO = Getcad("+", 1, AUXCAD)
        AUXCAD = Right(CAD, Len(CAD) - (Len(Getcad("+", j, CAD)) + 1))
        Set RsTmp = New ADODB.Recordset
        RsTmp.Open "SELECT " & CODCONCEPTO & " AS MONTO FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE CODTRAB='" & CODTRAB & "'", DBSYSTEM
        MONTO = ESNULO(RsTmp!MONTO, 0)
        TOTALGEN = 0
        If MONTO > 0 Then
            CONT = 0
            For I = 1 To 5
                MESANO = Format(DateAdd("m", -I, FECHA), "MMYYYY")
                STRCADENA = "SELECT SUM(MONTO) AS TOTAL FROM BOL" & MESANO & " B ,MOV" & MESANO & " M " & _
                         "WHERE B.INUMBOL=M.INUMBOL AND B.CODTRAB='" & CODTRAB & "' AND CONCEPTO='" & Trim(CODCONCEPTO) & "'"
                TOTAL = ESNULO(DevuelveValor(STRCADENA, DBSYSTEM), 0)
                If TOTAL > 0 And CONT < 2 Then
                    CONT = CONT + 1
                    TOTALGEN = TOTALGEN + TOTAL
                  ElseIf CONT = 2 Then Exit For
                End If
            Next
            If CONT = 2 Then
                TOTALGEN = (TOTALGEN + MONTO) / 3
              Else: TOTALGEN = 0
            End If
         End If
        MOTOGEN = MOTOGEN + TOTALGEN
    Next
    GET_HISTVAR = MOTOGEN
End Function

Private Sub AL_ACTUALIZAR_RECORSET_ANTERIOR() '(NO SE USA )
        'CALCULO DE FORMULAS TIPO INGRESO
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO <= 1 And Not !ESESCRITO And Not IsNull(!FORMULA) Then 'SI TIPO ES INGRESO
                DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=" & !FORMULA & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'" & IIf(Len(Trim(!CRITERIO)) = 0, "", " AND " & !CRITERIO)
                If Len(Trim(!CRITERIO)) = 0 Then DBSTARPLAN.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET " & !Codigo & "=0 WHERE CODTRAB='" & RSINPUT!CODTRAB & "'" & IIf(Len(Trim(!CRITERIO)) = 0, "", " AND NOT (" & !CRITERIO & ")")
            End If
            .MoveNext
        Loop
    End With
    DBSYSTEM.Execute STRCALCSUM & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
    
    '--------------------------------------
    'Calculo del Impuesto de 5ta. Categoria
    '--------------------------------------
    
Dim TOTALPERCIBIDO
Dim PROYECTADOFINAL
Dim TOTALRENTAPERCIBIR
Dim RENTAAFECTA
Dim IMPUESTOANUAL
Dim ACUMULADO
Dim SALDO
Dim MONTORETENER
Dim RETENCIONANTERIOR
    
    Dim XNUMM As Byte, VINGMIN As Double, M1 As Double, M2 As Double, S1 As Double, S2 As Double, M3 As Double, M4 As Double
    Dim RSAUX1 As New ADODB.Recordset, RSCALC As New ADODB.Recordset, RENTADEDUC As Double, XVALMES As Integer
    Dim RSTRAB As New ADODB.Recordset, XVALOPCION As Double, XC As Double
    Dim MESACUMULADO As Integer
    XNUMM = Month(REGINPUT.MESACTIVO) ' NUMERO DEL MES
    'AHORA USAMOS RSAUX1 PARA ABRIR EL RECORDSET DE IMPUESTOS ANTERIORES
    Set RSAUX1 = New ADODB.Recordset
    RSAUX1.Open "CONFIG5TA", DBSYSTEM, adOpenStatic
    Set RSCALC = New ADODB.Recordset
    RSCALC.Open "SELECT * FROM [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  WHERE AFECTOQUINTA=-1 AND CODTRAB='" & RSINPUT!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic 'SOLO AQUELLOS QUE ESTEN AFECTOS A QUINTA
    RENTADEDUC = RSAUX1!NUMUIT * RSAUX1!VALORUIT 'AHORA SON 7 UIT * 3000
    XVALMES = Val(RSAUX1("MES" & Format(XNUMM, "00"))) 'NUMERO QUE DIVIDIRA
    Select Case XNUMM 'SIRVE PARA SACAR EL ACUMULADO DE MESES ANTERIORES
        Case 1
            MESACUMULADO = RSAUX1!ACUMULA01
        Case 2
            MESACUMULADO = RSAUX1!ACUMULA02
        Case 3
            MESACUMULADO = RSAUX1!ACUMULA03
        Case 4
            MESACUMULADO = RSAUX1!ACUMULA04
        Case 5
            MESACUMULADO = RSAUX1!ACUMULA05
        Case 6
            MESACUMULADO = RSAUX1!ACUMULA06
        Case 7
            MESACUMULADO = RSAUX1!ACUMULA07
        Case 8
            MESACUMULADO = RSAUX1!ACUMULA08
        Case 9
            MESACUMULADO = RSAUX1!ACUMULA09
        Case 10
            MESACUMULADO = RSAUX1!ACUMULA10
        Case 11
            MESACUMULADO = RSAUX1!ACUMULA11
        Case 12
            MESACUMULADO = RSAUX1!ACUMULA12
    End Select
    XVALOPCION = 14 - Month(REGINPUT.MESACTIVO) + 1 'MES PROYECTADA
    Dim L As Single
    Do While Not RSCALC.EOF
        If ExisteTablaAux(" [##TMPQUINTA" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DELETE FROM  [##TMPQUINTA" & VGL_COMPUTER & "]  WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
        Set RSTRAB = New ADODB.Recordset
        RSTRAB.Open "SELECT TOTALEXTRA FROM TRABAJADORES WHERE CODTRAB='" & RSINPUT!CODTRAB & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
        If RSTRAB.RecordCount > 0 Then
            M1 = ((RSCALC!SUMARENTA + RSTRAB.Fields(0)) * XVALOPCION) + ((RSCALC!SUMARENTA + RSTRAB.Fields(0)) * (XNUMM - 1))  'X
            L = (M1 + RSCALC!T4 + RSCALC!T5) - RENTADEDUC
            If (L) > (54 * RSAUX1!VALORUIT) Then 'SI ES MAYOR A 54 UIT'S
                M3 = (54 * RSAUX1!VALORUIT) * (RSAUX1!PORCENTAJE / 100)
                M4 = (L - (54 * RSAUX1!VALORUIT)) * (RSAUX1!PORCENTAJE2 / 100)
                M2 = ((M3 + M4) - IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE MES>=" & (XNUMM - MESACUMULADO) & " AND MES<=" & (XNUMM - 1) & "  AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))) / XVALMES
                TOTALPERCIBIDO = (RSCALC!SUMARENTA)
                PROYECTADOFINAL = M1
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4 + RSCALC!T5)
                RENTAAFECTA = (TOTALRENTAPERCIBIR) - RENTADEDUC
                IMPUESTOANUAL = (M3 + M4)
                ACUMULADO = IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))
                SALDO = IMPUESTOANUAL - ACUMULADO
                MONTORETENER = M2
                RETENCIONANTERIOR = IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=" & (M2) & " WHERE CODTRAB='" & RSCALC!CODTRAB & "'"
            ElseIf (L) > (RENTADEDUC) Then 'SI ES MAYOR A 7 UIT'S
                M3 = (54 * RSAUX1!VALORUIT) * (RSAUX1!PORCENTAJE / 100)
                M2 = ((L) * (RSAUX1!PORCENTAJE / 100) - IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE MES>=" & (XNUMM - MESACUMULADO) & " AND MES<=" & (XNUMM - 1) & "  AND ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))) / XVALMES        '- RSTRAB!ACUMULAQUINTA
                TOTALPERCIBIDO = RSCALC!SUMARENTA
                PROYECTADOFINAL = M1
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4 + RSCALC!T5)
                RENTAAFECTA = (TOTALRENTAPERCIBIR) - RENTADEDUC
                IMPUESTOANUAL = (RENTAAFECTA) * (RSAUX1!PORCENTAJE / 100)
                ACUMULADO = IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))
                SALDO = IMPUESTOANUAL - ACUMULADO
                MONTORETENER = M2
                RETENCIONANTERIOR = IIf(MESACUMULADO = 0, 0, Fix(DevuelveValor("SELECT SUM([MONTO RETENER]) FROM HIST5TA WHERE HIST5TA.MES>=" & (XNUMM - MESACUMULADO) & " AND HIST5TA.MES<=" & (XNUMM - 1) & "  AND HIST5TA.ANNO='" & Year(REGINPUT.MESACTIVO) & "' AND HIST5TA.CODTRAB='" & RSCALC!CODTRAB & "'", DBSYSTEM)))
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=" & (M2) & " WHERE CODTRAB='" & RSCALC!CODTRAB & "'"
            Else
                TOTALPERCIBIDO = RSCALC!SUMARENTA
                PROYECTADOFINAL = M1
                TOTALRENTAPERCIBIR = (M1 + RSCALC!T4 + RSCALC!T5)
                RENTAAFECTA = 0
                IMPUESTOANUAL = 0
                ACUMULADO = 0
                SALDO = 0
                MONTORETENER = 0
                RETENCIONANTERIOR = 0
                DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET [_QUINTACAT]=0" & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
            End If
                'GRABA EN UN TEMPORAL LA TABLA DETALLADA DE LA QUINTA
                DBSYSTEM.Execute "INSERT INTO  [##TMPQUINTA" & VGL_COMPUTER & "]  (MES , ANNO , CODTRAB , NOMBRES , [TOTAL PERCIBIDO] , [PROYECTADO FIN AÑO] , [TOTAL RENTA PERCIBIR], [RENTA AFECTA] , [IMPUESTO ANUAL] , [ACUMULADO] , SALDO , [MONTO RETENER] , [RENTENCION ANTERIOR] ) VALUES (" & _
                "" & Month(REGINPUT.MESACTIVO) & ",'" & Year(REGINPUT.MESACTIVO) & "','" & RSINPUT!CODTRAB & "','" & RSCALC!NOMBRES & "'," & TOTALPERCIBIDO & "," & PROYECTADOFINAL & _
                "," & TOTALRENTAPERCIBIR & "," & RENTAAFECTA & "," & IMPUESTOANUAL & "," & ACUMULADO & "," & SALDO & "," & _
                MONTORETENER & "," & RETENCIONANTERIOR & ")"
                
        End If
        Set RSTRAB = Nothing
        RSCALC.MoveNext
    Loop
    Set RSAUX1 = Nothing
    Set RSAUX = Nothing
    Set RSCALC = Nothing
    
    
    
    'CALCULO DEL TOTAL EGRESOS
    '-------------------------
    'YA SE INCLUYEN EN ESTE PROCESO LOS OTROS EGRESOS
    Dim CADSUMAS As String
    CADSUMAS = "0+OTROSEGR+ADELANTO"
    With RSCNPT
        .MoveFirst
        Do While Not .EOF
            If !TIPO = 2 Then
                'PARA EL TOTAL DE INGRESOS
                CADSUMAS = CADSUMAS & "+" & !Codigo
            End If
            .MoveNext
        Loop
        DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET TOTEGR=" & CADSUMAS & " WHERE CODTRAB='" & RSINPUT!CODTRAB & "'"
    End With
    If REGINPUT.REDONDEO Then
        DBSYSTEM.Execute "UPDATE [##CALCINPUT" & Trim(VGL_COMPUTER) & "]  SET XREDONDEO=CASE WHEN ([TOTING]-[TOTEGR])-ROUND([TOTING]-[TOTEGR],0,1)>=0.5 THEN " _
            & "1-(([TOTING]-[TOTEGR])-(ROUND([TOTING]-[TOTEGR],0,1))) ELSE -1*(([TOTING]-[TOTEGR])-(ROUND([TOTING]-[TOTEGR],0,1))) END"
    End If
    CALCULOTOTALGEN
End Sub
Private Function RPLCADENA(ByVal CADENA As String, ByVal CODTRAB As String, FECHA As Date, NOMBOL As Long) As String
    Dim POSARROBA As Integer, POS1 As Integer, PROCESO As String, CAMPO As String, POS2 As Integer
    Dim VALOR As Double
    POSARROBA = 1
    POSARROBA = InStr(POSARROBA, CADENA, "@")
    Do While POSARROBA <> 0
        POS1 = InStr(POSARROBA, CADENA, "(")
        PROCESO = Mid(CADENA, POSARROBA + 1, POS1 - (POSARROBA + 1))
        POS2 = InStr(POSARROBA, CADENA, ")")
        CAMPO = Mid(CADENA, POS1 + 1, POS2 - (POS1 + 1))
        VALOR = CALCCNP(CODTRAB, CAMPO, FECHA, NOMBOL, PROCESO)
        If IsNull(VALOR) Then VALOR = 0
        CADENA = Replace(CADENA, Mid(CADENA, POSARROBA, (POS2 - POSARROBA) + 1), "" & VALOR)
        POSARROBA = InStr(POSARROBA, CADENA, "@")
    Loop
   RPLCADENA = CADENA
End Function
Private Function CALCCNP(CODTRAB As String, CONCEPTO As String, FECHA As Date, NOMBOL As Long, PROCESO As String) As Double
    Dim PERIODO As String
    Dim ULTCRON As Boolean
    PERIODO = Format(FECHA, "mmyyyy")
    CALCCNP = 0
    Select Case UCase(PROCESO)
        Case "VARIABLE"
            ULTCRON = GetValor("SELECT ULTMES FROM NOMBOL WHERE CODIGO=" & NOMBOL, DBSYSTEM)
            If Not ULTCRON Then Exit Function
            CALCCNP = Round(ESNULO(GetValor("SELECT SUM(" & CONCEPTO & ") FROM BOL" & PERIODO & " WHERE CODTRAB='" & CODTRAB & "'", DBSYSTEM), 0), 2)
        Case "SUMAVAR"
            CALCCNP = Round(ESNULO(GetValor("SELECT SUM(" & CONCEPTO & ") FROM BOL" & PERIODO & " WHERE CODTRAB='" & CODTRAB & "'", DBSYSTEM), 0), 2)
    End Select
End Function
Private Sub GRABAR_EN_INGMOV2000()
  Dim RSayuda As ADODB.Recordset
  Dim K As Integer
  Dim RSINGMOV200 As ADODB.Recordset
  Dim CAD As String
  
  Set RSayuda = New ADODB.Recordset
  RSayuda.Open "[##CALCINPUT" & Trim(VGL_COMPUTER) & "] ", DBSYSTEM, adOpenDynamic, adLockReadOnly
  
  
  
  'ingmov2000
  'CODTRAB CONCEPTO   CODNOMBOL  ----VALOR
  
Do While Not RSayuda.EOF
    For K = 2 To dgInput.Columns.Count - 1
        'VERIFICAR SI ES DE TIPO ESCRITO
     CAD = "SELECT TIPO FROM CONCEPTOS WHERE CODIGO='" & dgInput.Columns(K).Caption & "'"
     If (DevuelveValor(CAD, DBSYSTEM) = 1) Or (DevuelveValor(CAD, DBSYSTEM) = 2) Then
         CAD = "SELECT ESESCRITO FROM CONCEPTOS WHERE CODIGO='" & dgInput.Columns(K).Caption & "'"
         If DevuelveValor(CAD, DBSYSTEM) = True Then
             If RSayuda.Fields(dgInput.Columns(K).Caption) > 0 Then
                '******************************************************
                Set RSINGMOV200 = New ADODB.Recordset
                RSINGMOV200.Open "SELECT * FROM INGMOV2000  WHERE  CODTRAB='" & RSayuda.Fields("CODTRAB") & "'  AND  CONCEPTO='" & dgInput.Columns(K).Caption & "' AND CODNOMBOL=" & REGINPUT.Codigo & "", DBSYSTEM, adOpenDynamic, adLockOptimistic
                If RSINGMOV200.EOF Then
                    RSINGMOV200.AddNew
                        RSINGMOV200!CODTRAB = RSayuda.Fields("CODTRAB")
                        RSINGMOV200!CONCEPTO = dgInput.Columns(K).Caption
                        RSINGMOV200!CODNOMBOL = REGINPUT.Codigo
                        RSINGMOV200!VALOR = RSayuda.Fields(dgInput.Columns(K).Caption)
                    RSINGMOV200.Update
                Else
                        RSINGMOV200!VALOR = RSayuda.Fields(dgInput.Columns(K).Caption)
                    RSINGMOV200.Update
                End If
                '********************************************************
             End If
        End If
    End If
 Next K
 RSayuda.MoveNext
Loop

End Sub
