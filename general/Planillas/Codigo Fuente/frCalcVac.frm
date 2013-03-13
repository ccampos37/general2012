VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frCalcVac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Vacaciones"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frCalcVac.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cálculo nuevo de Vacaciones"
      Height          =   4350
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   7815
      Begin TabDlg.SSTab ssTab1 
         Height          =   2745
         Left            =   180
         TabIndex        =   13
         Top             =   1455
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   4842
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Calculo General"
         TabPicture(0)   =   "frCalcVac.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "xTotal"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "dgProceso"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Command1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmCancelar"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmAceptar"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "xCalculo"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Cálculo Detallado"
         TabPicture(1)   =   "frCalcVac.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgDetalle"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Mensajes del Proceso"
         TabPicture(2)   =   "frCalcVac.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mensajes"
         Tab(2).ControlCount=   1
         Begin AplisetControlText.Aplitext xCalculo 
            Height          =   315
            Left            =   5760
            TabIndex        =   2
            Top             =   1770
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            Text            =   "0.00"
         End
         Begin VB.CommandButton cmAceptar 
            Caption         =   "&Aceptar"
            Height          =   360
            Left            =   5970
            TabIndex        =   20
            Top             =   375
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.ListBox Mensajes 
            Height          =   2010
            Left            =   -74865
            TabIndex        =   18
            Top             =   435
            Width           =   7155
         End
         Begin MSDataGridLib.DataGrid dgDetalle 
            Height          =   2100
            Left            =   -74865
            TabIndex        =   17
            Top             =   480
            Visible         =   0   'False
            Width           =   5955
            _ExtentX        =   10504
            _ExtentY        =   3704
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
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.CommandButton cmCancelar 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            Height          =   360
            Left            =   5970
            TabIndex        =   16
            Top             =   855
            Visible         =   0   'False
            Width           =   1410
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Calcular"
            Height          =   360
            Left            =   5970
            TabIndex        =   15
            Top             =   375
            Width           =   1410
         End
         Begin MSDataGridLib.DataGrid dgProceso 
            Height          =   2295
            Left            =   195
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   5040
            _ExtentX        =   8890
            _ExtentY        =   4048
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
               MarqueeStyle    =   4
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label xTotal 
            AutoSize        =   -1  'True
            Caption         =   "Remuneracion Vacacional"
            Height          =   195
            Left            =   5505
            TabIndex        =   19
            Top             =   1440
            Width           =   1875
         End
      End
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   300
         Left            =   4035
         TabIndex        =   8
         Top             =   1005
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36709
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   300
         Left            =   2100
         TabIndex        =   6
         Top             =   1005
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36709
      End
      Begin MSComCtl2.DTPicker xFechaIng 
         Height          =   300
         Left            =   165
         TabIndex        =   4
         Top             =   1005
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36709
      End
      Begin VB.Label xBasico 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00 "
         Height          =   300
         Left            =   6480
         TabIndex        =   12
         Top             =   345
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Basico"
         Height          =   195
         Left            =   5880
         TabIndex        =   11
         Top             =   375
         Width           =   480
      End
      Begin VB.Label xArea 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   5880
         TabIndex        =   10
         Top             =   1005
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Area"
         Height          =   195
         Left            =   5940
         TabIndex        =   9
         Top             =   795
         Width           =   330
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   4035
         TabIndex        =   7
         Top             =   795
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Left            =   2100
         TabIndex        =   5
         Top             =   780
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Ingreso"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cálculo de Vacaciones"
      Height          =   4350
      Left            =   90
      TabIndex        =   21
      Top             =   75
      Width           =   7815
      Begin VB.CommandButton Command2 
         Caption         =   "&Cerrar"
         Height          =   405
         Left            =   5985
         TabIndex        =   25
         Top             =   1365
         Width           =   1485
      End
      Begin VB.CommandButton cmNuevo 
         Caption         =   "&Nuevo Calculo"
         Height          =   405
         Left            =   5985
         TabIndex        =   23
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton cmEliminar 
         Caption         =   "&Eliminar Calculo"
         Height          =   405
         Left            =   5985
         TabIndex        =   22
         Top             =   810
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid dgLista 
         Height          =   3945
         Left            =   255
         TabIndex        =   24
         Top             =   240
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6959
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         Caption         =   "Calculos ya realizados y pendientes de transferir a Planilla"
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
            MarqueeStyle    =   4
            AllowRowSizing  =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "frCalcVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSAUX As New ADODB.Recordset
Dim RSLISTA As New ADODB.Recordset
Dim RSPROC As New ADODB.Recordset
Dim RSDETS As New ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    DBSYSTEM.Execute "UPDATE HISTOVAC SET MONTO=" & xCalculo.Text & " WHERE CODIGO=" & xTotal.Tag
    REFRESCARDG
    Frame1.Visible = False
    Frame2.Visible = True
End Sub

Private Sub CMCANCELAR_CLICK()
    Frame1.Visible = False
    Frame2.Visible = True
    cmCancelar.Visible = False
    Mensajes.Clear
    dgProceso.Visible = False
    dgDetalle.Visible = False
    cmAceptar.Visible = False
    Command1.Visible = True
End Sub

Private Sub CMELIMINAR_CLICK()
    If RSLISTA.EOF Or RSLISTA.RecordCount = 0 Then
        MsgBox "ERROR DE USUARIO: NO EXISTE REGISTRO PARA ELIMINAR EL CALCULO DE VACACIONES", vbCritical
        Exit Sub
    End If
    If MsgBox("RELAMENTE DESEA ELIMINAR EL CÁLCULO DEL TRABAJADOR " & RSLISTA!NOMBRES, vbYesNo + vbInformation) = vbNo Then Exit Sub
    DBSYSTEM.Execute "UPDATE HISTOVAC SET MONTO=0 WHERE CODIGO=" & RSLISTA!Codigo
    REFRESCARDG
End Sub

Private Sub CMNUEVO_CLICK()
    Set RSAUX = Nothing
    RSAUX.Open "SELECT VWTRABAJ.CODTRAB, NOMBRES, NOMBREAREA, VWTRABAJ.FECHAING, VWTRABAJ.BASICO, FECHAINICAL, FECHAFINCAL,CODIGO FROM VWTRABAJ,HISTOVAC WHERE VWTRABAJ.CODTRAB=HISTOVAC.CODTRAB AND CERRADO<>1 AND (MONTO=0 OR (MONTO)IS NULL)", DBSYSTEM, adOpenStatic
    If RSAUX.RecordCount = 0 Then
        MsgBox "NO EXISTEN TRABAJADORES PROGRAMADOS PARA VACACIONES", vbCritical
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX, , "VWTRABAJ.CODTRAB"
    frmComun.Show 1
    If VGUTIL(1) = "" Then
        Exit Sub
    End If
    xTrab.Text = RSAUX!NOMBRES
    xTrab.Tag = RSAUX!CODTRAB
    xTotal.Tag = RSAUX!Codigo
    xFechaIng.Value = RSAUX!FECHAING
    xFechaIni.Value = RSAUX!FECHAINICAL
    xFechaFin.Value = RSAUX!FECHAFINCAL
    xBasico.Caption = Format(RSAUX!BASICO, "0.00 ")
    xArea.Caption = RSAUX!NOMBREAREA & " "
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub Command1_Click()
    CALCULARVAC
    SSTab1.Tab = 2
    Command1.Visible = False
    cmAceptar.Visible = True
    cmCancelar.Visible = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Frame1.Visible = False
    RSLISTA.Open "SELECT NOMBRES, HISTOVAC.* FROM HISTOVAC, VWTRABAJ WHERE HISTOVAC.CODTRAB=VWTRABAJ.CODTRAB AND MONTO<>0 AND CERRADO<>1", DBSYSTEM, adOpenStatic
    REFRESCARDG
End Sub

Public Sub REFRESCARDG()
    RSLISTA.Requery
    Set DGLista.DataSource = RSLISTA
    With DGLista
        .Columns("CODIGO").Visible = False
        .Columns("CODTRAB").Visible = False
        .Columns("FECHAING").Visible = False
        .Columns("AREA").Visible = False
        .Columns("BASICO").Visible = False
        .Columns("DIAS").Visible = False
        .Columns("FECHAREG").Visible = False
        .Columns("CERRADO").Visible = False
        .Columns("MONTO").Alignment = dbgRight
        .Columns("MONTO").NumberFormat = "##,##0.00 "
        .Columns("NOMBRES").Width = 2063
    End With
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSLISTA = Nothing
    Set RSAUX = Nothing
End Sub

Public Sub CALCULARVAC()
    If xTrab.Tag = "" Then
        MsgBox "FALTA SELECCIONAR UN TRABAJADOR", vbCritical
        Exit Sub
    End If
    If xFechaIni.Value < xFechaIng.Value Then
        MsgBox "EXISTE UN ERROR EN EL RANGO DE FECHAS PARA EL CÁLCULO DE VACACIONES. LA FECHA DE INICIO DEL CÁLCULO NO PUEDE SER MENOR A LA FECHA DE INGRESO DEL TRABAJADOR", vbCritical
        Exit Sub
    End If
    If ExisteTablaAux("CALCVAC") Then DBSYSTEM.Execute "DROP TABLE CALCVAC"
    DBSYSTEM.Execute "CREATE TABLE CALCVAC (CONCEPTO VARCHAR(50), TIPO BIT, MES VARCHAR(6), MONTO  Numeric(20,2) )"
    Dim RSNOMBOL As ADODB.Recordset
    Dim STRMES As String
    Set RSNOMBOL = New ADODB.Recordset
    RSNOMBOL.Open "SELECT * FROM NOMBOL WHERE FECHAINI>=" & DateSQL(xFechaIni.Value) & " AND FECHAINI<" & DateSQL(xFechaFin.Value), DBSYSTEM, adOpenStatic
    If RSNOMBOL.RecordCount = 0 Then
        Mensajes.AddItem "NO SE HAN ENCONTRADO BOLETAS DE REMUNERACIONES DENTRO DEL RANGO DE FECHAS ESPECIFICADO"
    End If
    Mensajes.Clear
    'RECICLAMOS EL USO DE REGINPUT
    Do While Not RSNOMBOL.EOF
        REGINPUT.BOL_TABLE = "BOL" & Format(Month(RSNOMBOL!MES), "00") & Year(RSNOMBOL!MES)
        REGINPUT.MOV_TABLE = "MOV" & Format(Month(RSNOMBOL!MES), "00") & Year(RSNOMBOL!MES)
        REGINPUT.Codigo = RSNOMBOL!Codigo
        If Not ExisteTabla(REGINPUT.BOL_TABLE) Then
            Mensajes.AddItem "    ** NO SE HAN DEFINIDO PAGOS DEL MES DE " & AMESES(Month(RSNOMBOL!MES)) & " DE " & Year(RSNOMBOL!MES)
        Else
            STRMES = Year(RSNOMBOL!MES) & Format(Month(RSNOMBOL!MES), "00")
            Set RSAUX = Nothing
            RSAUX.Open "SELECT CONCEPTOS.NOMBRE,CONCEPTOS.TIPOREMU,MOV.MONTO FROM " & REGINPUT.BOL_TABLE & " BOL," & REGINPUT.MOV_TABLE & " MOV, CONCEPTOS WHERE BOL.INUMBOL=MOV.INUMBOL AND MOV.CONCEPTO=CONCEPTOS.CODIGO " & " AND CONCEPTOS.TIPO=1 AND CONCEPTOS.TIPOREMU<>2 AND BOL.CODTRAB='" & xTrab.Tag & "' AND BOL.CODNOMBOL=" & REGINPUT.Codigo, DBSYSTEM, adOpenStatic
            If RSAUX.RecordCount = 0 Then
                Mensajes.AddItem "         ** NO PRESENTA BOLETAS EN: " & RSNOMBOL!NOMBRE
            Else
                Mensajes.AddItem "CARGA DE BOLETA DE REM. DE " & RSNOMBOL!NOMBRE
                Do While Not RSAUX.EOF
                    DBSYSTEM.Execute "INSERT INTO CALCVAC VALUES ('" & RSAUX!NOMBRE & "'," & RSAUX!TIPOREMU & ",'" & STRMES & "'," & RSAUX!MONTO & ")"
                    RSAUX.MoveNext
                Loop
            End If
        End If
        RSNOMBOL.MoveNext
    Loop
    Set RSDETS = Nothing
    RSDETS.Open "CALCVAC", DBSYSTEM, adOpenStatic
    Set dgDetalle.DataSource = RSDETS
    Set RSPROC = Nothing
    RSPROC.Open "SELECT CONCEPTO, SUM(MONTO) AS MONTOCALC FROM CALCVAC GROUP BY CONCEPTO", DBSYSTEM, adOpenStatic
    Set dgProceso.DataSource = RSPROC
    Set RSAUX = Nothing
    RSAUX.Open "SELECT SUM(MONTO) AS TOTAL FROM CALCVAC WHERE TIPO=0", DBSYSTEM, adOpenStatic
    Dim XVAL As Single
    XVAL = IIf(IsNull(RSAUX!TOTAL), 0, RSAUX!TOTAL)
    Set RSAUX = Nothing
    RSAUX.Open "SELECT SUM(MONTO) AS TOTAL FROM CALCVAC WHERE TIPO=1", DBSYSTEM, adOpenStatic
    XVAL = (XVAL + IIf(IsNull(RSAUX!TOTAL), 0, RSAUX!TOTAL) / 6) / 6
    xCalculo.Text = Format(XVAL, "0.00 ")
    dgProceso.Columns("MONTOCALC").Alignment = dbgRight
    dgProceso.Columns("MONTOCALC").NumberFormat = "##,##0.00 "
    dgProceso.Columns("CONCEPTO").Width = 2954.835
    dgDetalle.Columns("MONTO").Alignment = dbgRight
    dgDetalle.Columns("MONTO").NumberFormat = "##,##.00 "
    dgDetalle.Columns("CONCEPTO").Width = 3314.835
    dgProceso.Visible = True
    dgDetalle.Visible = True
    Set RSNOMBOL = Nothing
    cmAceptar.Enabled = True
End Sub

