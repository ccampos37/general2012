VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frCalcGrat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo de Gratificaciones"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   Icon            =   "frCalcGrat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Cálculo Nuevo de Gratificación"
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
         TabPicture(0)   =   "frCalcGrat.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "xTotal"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "xCalculo"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "dgProceso"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Command1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmCancelar"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmAceptar"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Cálculo Detallado"
         TabPicture(1)   =   "frCalcGrat.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dgDetalle"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Mensajes del Proceso"
         TabPicture(2)   =   "frCalcGrat.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Mensajes"
         Tab(2).ControlCount=   1
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
         Begin VB.Label xCalculo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Sin Calculo "
            Height          =   330
            Left            =   5520
            TabIndex        =   21
            Top             =   1725
            Width           =   1875
         End
         Begin VB.Label xTotal 
            AutoSize        =   -1  'True
            Caption         =   "Gratificación Ordinaria"
            Height          =   195
            Left            =   5505
            TabIndex        =   19
            Top             =   1440
            Width           =   1560
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
         Format          =   23592961
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
         Format          =   23592961
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
         Format          =   23592961
         CurrentDate     =   36709
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   285
         Left            =   1125
         TabIndex        =   2
         Top             =   360
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         Text            =   ""
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
      TabIndex        =   22
      Top             =   75
      Width           =   7815
      Begin VB.ComboBox xMes 
         Height          =   315
         ItemData        =   "frCalcGrat.frx":091E
         Left            =   2775
         List            =   "frCalcGrat.frx":0928
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   315
         Width           =   2985
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cerrar"
         Height          =   405
         Left            =   5985
         TabIndex        =   26
         Top             =   1365
         Width           =   1485
      End
      Begin VB.CommandButton cmNuevo 
         Caption         =   "&Nuevo Calculo"
         Height          =   405
         Left            =   5985
         TabIndex        =   24
         Top             =   240
         Width           =   1485
      End
      Begin VB.CommandButton cmEliminar 
         Caption         =   "&Eliminar Calculo"
         Height          =   405
         Left            =   5985
         TabIndex        =   23
         Top             =   810
         Width           =   1485
      End
      Begin MSDataGridLib.DataGrid dgLista 
         Height          =   3435
         Left            =   255
         TabIndex        =   25
         Top             =   750
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6059
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gratificación correspondiente a"
         Height          =   195
         Left            =   285
         TabIndex        =   28
         Top             =   360
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frCalcGrat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsAux As New ADODB.Recordset
Dim RsLista As New ADODB.Recordset
Dim RsProc As New ADODB.Recordset
Dim RsDets As New ADODB.Recordset

Private Sub cmAceptar_Click()
    DbSystem.Execute "UPDATE HistoVac SET Monto=" & xCalculo.Caption & " WHERE Codigo=" & xTotal.Tag
    RefrescarDG
    Frame1.Visible = False
    Frame2.Visible = True
End Sub

Private Sub cmCancelar_Click()
    Frame1.Visible = False
    Frame2.Visible = True
    cmCancelar.Visible = False
    Mensajes.Clear
    dgProceso.Visible = False
    dgDetalle.Visible = False
    cmAceptar.Visible = False
    Command1.Visible = True
End Sub

Private Sub cmEliminar_Click()
    If RsLista.EOF Or RsLista.RecordCount = 0 Then
        MsgBox "Error de Usuario: No existe registro para eliminar el Calculo de Vacaciones", vbCritical
        Exit Sub
    End If
    If MsgBox("Relamente desea eliminar el Cálculo del Trabajador " & RsLista!Nombres, vbYesNo + vbInformation) = vbNo Then Exit Sub
    DbSystem.Execute "UPDATE HistoVac SET Monto=0 WHERE Codigo=" & RsLista!codigo
    RefrescarDG
End Sub

Private Sub cmNuevo_Click()
    Set RsAux = Nothing
    RsAux.Open "SELECT vwTrabaj.CodTrab, Nombres, NombreArea, vwTrabaj.FechaIng, vwTrabaj.Basico, FechaIniCal, FechaFinCal,Codigo FROM vwTrabaj,HistoVac WHERE vwTrabaj.CodTrab=HistoVac.CodTrab AND Cerrado<>1 AND (Monto=0 OR IsNull(Monto))", DbSystem, adOpenStatic
    If RsAux.RecordCount = 0 Then
        MsgBox "No existen trabajadores programados para Vacaciones", vbCritical
        Exit Sub
    End If
    frmComun.Conectar RsAux, , "vwTrabaj.CodTrab"
    frmComun.Show 1
    If vgUtil(1) = "" Then
        Exit Sub
    End If
    xTrab.Text = RsAux!Nombres
    xTrab.Tag = RsAux!CodTrab
    xTotal.Tag = RsAux!codigo
    xFechaIng.Value = RsAux!FechaIng
    xFechaIni.Value = RsAux!FechaIniCal
    xFechaFin.Value = RsAux!FechaFinCal
    xBasico.Caption = Format(RsAux!Basico, "0.00 ")
    xArea.Caption = RsAux!NombreArea & " "
    Frame1.Visible = True
    Frame2.Visible = False
End Sub

Private Sub Command1_Click()
    CalcularVac
    SSTab1.Tab = 2
    Command1.Visible = False
    cmAceptar.Visible = True
    cmCancelar.Visible = True
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    xMes.ListIndex = 1
    Frame1.Visible = False
    RsLista.Open "SELECT Nombres, HistoVac.* FROM HistoVac, vwTrabaj WHERE HistoVac.CodTrab=vwTrabaj.CodTrab AND Monto<>0 AND Cerrado<>1", DbSystem, adOpenStatic
    RefrescarDG
End Sub

Public Sub RefrescarDG()
    RsLista.Requery
    Set DGLista.DataSource = RsLista
    With DGLista
        .Columns("Codigo").Visible = False
        .Columns("CodTrab").Visible = False
        .Columns("FechaIng").Visible = False
        .Columns("Area").Visible = False
        .Columns("Basico").Visible = False
        .Columns("Dias").Visible = False
        .Columns("FechaReg").Visible = False
        .Columns("Cerrado").Visible = False
        .Columns("monto").Alignment = dbgRight
        .Columns("monto").NumberFormat = "##,##0.00 "
        .Columns("nombres").Width = 2063
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsLista = Nothing
    Set RsAux = Nothing
End Sub

Public Sub CalcularVac()
    If xFechaIni.Value < xFechaIng.Value Then
        MsgBox "Existe un error en el rango de fechas para el cálculo de vacaciones. La fecha de Inicio del cálculo no puede ser menor a la fecha de Ingreso del trabajador", vbCritical
        Exit Sub
    End If
    If xTrab.Tag = "" Then
        MsgBox "Falta seleccionar un trabajador", vbCritical
        Exit Sub
    End If
    If ExisteTablaAux("CalcVac") Then DBAuxCom.Execute "DROP TABLE CalcVac"
    DBAuxCom.Execute "CREATE TABLE Calcvac (Concepto Text(50), Tipo Byte, Mes Text(6), Monto Single)"
    Dim RsNomBol As ADODB.Recordset
    Dim strMes As String
    Set RsNomBol = New ADODB.Recordset
    RsNomBol.Open "SELECT * FROM NomBol WHERE FechaIni>=" & DateSQL(xFechaIni.Value) & " AND FechaIni<" & DateSQL(xFechaFin.Value), DbSystem, adOpenStatic
    If RsNomBol.RecordCount = 0 Then
        Mensajes.AddItem "No se han encontrado boletas de remuneraciones dentro del rango de fechas especificado"
    End If
    Mensajes.Clear
    'Reciclamos el uso de RegInput
    Do While Not RsNomBol.EOF
        RegInput.Bol_Table = "BOL" & Format(Month(RsNomBol!Mes), "00") & Year(RsNomBol!Mes)
        RegInput.Mov_Table = "MOV" & Format(Month(RsNomBol!Mes), "00") & Year(RsNomBol!Mes)
        RegInput.codigo = RsNomBol!codigo
        If Not ExisteTabla(RegInput.Bol_Table) Then
            Mensajes.AddItem "    ** No se han definido pagos del mes de " & AMeses(Month(RsNomBol!Mes)) & " de " & Year(RsNomBol!Mes)
        Else
            strMes = Year(RsNomBol!Mes) & Format(Month(RsNomBol!Mes), "00")
            Set RsAux = Nothing
            RsAux.Open "SELECT Conceptos.Nombre,Conceptos.TipoRemu,Mov.Monto FROM " & RegInput.Bol_Table & " BOL," & RegInput.Mov_Table & " MOV, Conceptos WHERE Bol.INumBol=Mov.INumBol AND Mov.Concepto=Conceptos.Codigo " & " AND Conceptos.Tipo=1 AND Conceptos.TipoRemu<>2 AND Bol.CodTrab='" & xTrab.Tag & "' AND Bol.CodNomBol=" & RegInput.codigo, DbSystem, adOpenStatic
            If RsAux.RecordCount = 0 Then
                Mensajes.AddItem "         ** No presenta boletas en: " & RsNomBol!Nombre
            Else
                Mensajes.AddItem "Carga de Boleta de Rem. de " & RsNomBol!Nombre
                Do While Not RsAux.EOF
                    DBAuxCom.Execute "INSERT INTO CalcVac VALUES ('" & RsAux!Nombre & "'," & RsAux!TipoRemu & ",'" & strMes & "'," & RsAux!Monto & ")"
                    RsAux.MoveNext
                Loop
            End If
        End If
        RsNomBol.MoveNext
    Loop
    Set RsDets = Nothing
    RsDets.Open "CalcVac", DBAuxCom, adOpenStatic
    Set dgDetalle.DataSource = RsDets
    Set RsProc = Nothing
    RsProc.Open "SELECT Concepto, Sum(Monto) AS MontoCalc FROM CalcVac GROUP BY Concepto", DBAuxCom, adOpenStatic
    Set dgProceso.DataSource = RsProc
    Set RsAux = Nothing
    RsAux.Open "SELECT Sum(Monto) AS Total FROM CalcVac Where Tipo=0", DBAuxCom, adOpenStatic
    Dim xVal As Single
    xVal = RsAux!Total
    Set RsAux = Nothing
    RsAux.Open "SELECT Sum(Monto) AS Total FROM CalcVac Where Tipo=1", DBAuxCom, adOpenStatic
    xVal = (xVal + RsAux!Total / 6) / 6
    xCalculo.Caption = Format(xVal, "0.00 ")
    dgProceso.Columns("MontoCalc").Alignment = dbgRight
    dgProceso.Columns("MontoCalc").NumberFormat = "##,##0.00 "
    dgProceso.Columns("Concepto").Width = 2954.835
    dgDetalle.Columns("Monto").Alignment = dbgRight
    dgDetalle.Columns("Monto").NumberFormat = "##,##.00 "
    dgDetalle.Columns("Concepto").Width = 3314.835
    dgProceso.Visible = True
    dgDetalle.Visible = True
    Set RsNomBol = Nothing
    cmAceptar.Enabled = True
End Sub
