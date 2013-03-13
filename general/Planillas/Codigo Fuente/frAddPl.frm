VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#7.0#0"; "ApliCTxt.ocx"
Begin VB.Form frAddPl 
   Caption         =   "Agregar Planilla"
   ClientHeight    =   6750
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6330
   Icon            =   "frAddPl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6750
   ScaleWidth      =   6330
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox xTipo 
      Height          =   315
      ItemData        =   "frAddPl.frx":030A
      Left            =   1950
      List            =   "frAddPl.frx":031D
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5775
      Width           =   2730
   End
   Begin AplisetControlText.Aplitext xCCosto 
      Height          =   285
      Left            =   1950
      TabIndex        =   13
      Top             =   4725
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   3390
      TabIndex        =   11
      Top             =   6300
      Width           =   1320
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   330
      Left            =   1620
      TabIndex        =   10
      Top             =   6315
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker xFechaFin 
      Height          =   300
      Left            =   4740
      TabIndex        =   9
      Top             =   5430
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24379393
      CurrentDate     =   36679
   End
   Begin MSComCtl2.DTPicker xFechaIni 
      Height          =   300
      Left            =   1950
      TabIndex        =   7
      Top             =   5415
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   529
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   24379393
      CurrentDate     =   36679
   End
   Begin AplisetControlText.Aplitext xNombre 
      Height          =   285
      Left            =   1950
      TabIndex        =   4
      Top             =   5070
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   503
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Planilla"
      Height          =   4335
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   6225
      Begin MSDataGridLib.DataGrid DGCrono 
         Height          =   3225
         Left            =   420
         TabIndex        =   5
         Top             =   585
         Width           =   5610
         _ExtentX        =   9895
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
         Caption         =   "CRONOGRAMA DE PAGOS"
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
            MarqueeStyle    =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton Op1 
         Caption         =   "De acuerdo a Cronograma de Pagos"
         Height          =   240
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Value           =   -1  'True
         Width           =   2940
      End
      Begin VB.OptionButton Op1 
         Caption         =   "Planilla Libre"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   3960
         Width           =   1245
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Planilla"
      Height          =   195
      Left            =   510
      TabIndex        =   14
      Top             =   5850
      Width           =   1080
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Centro de Costo"
      Height          =   195
      Left            =   510
      TabIndex        =   12
      Top             =   4770
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Término"
      Height          =   195
      Left            =   3360
      TabIndex        =   8
      Top             =   5475
      Width           =   1290
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Inicio"
      Height          =   195
      Left            =   510
      TabIndex        =   6
      Top             =   5475
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de Planilla"
      Height          =   195
      Left            =   510
      TabIndex        =   0
      Top             =   5115
      Width           =   1320
   End
End
Attribute VB_Name = "frAddPl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RsCrono As ADODB.Recordset
Attribute RsCrono.VB_VarHelpID = -1

Private Sub cmAceptar_Click()
    If Op1(0).Value Then      'De acuerdo a cronograma, tenemos que marcar para no editarlo nuevamente, ni eliminarlo desde su panel principal
        DbSystem.Execute "UPDATE FechaPago SET Cerrado=1 WHERE ID_FechaPago=" & RsCrono!ID_FechaPago
    End If
    DbSystem.Execute "INSERT INTO NomBol (TIPOPLANILLA, NOMBRE, CCOSTO, MES, FECHAINI, FECHAFIN) VALUES (" & xTipo.ListIndex & ",'" & xNombre.Text & "','" & xCCosto.Tag & "'," & DateSQL(vpFecha) & "," & DateSQL(xFechaIni.Value) & "," & DateSQL(xFechaFin.Value) & ")"
    Unload Me
End Sub

Private Sub cmCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim F1 As Date, F2 As Date
    F1 = CDate("01/" & Month(vpFecha) & "/" & Year(vpFecha))
    F2 = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Month(vpFecha) & "/" & Year(vpFecha))))
    Set RsCrono = New ADODB.Recordset
    RsCrono.Open "SELECT CCostos.CodCCosto, CCostos.Nombre, FechaPago.Nombre AS Descripcion, ID_FechaPago, FechaIni, FechaFin FROM CCostos, FechaPago WHERE CCostos.CodCCosto=FechaPago.CodCCosto AND Cerrado=0 AND (FechaIni Between " & DateSQL(F1) & " AND " & DateSQL(F2) & ") ORDER BY FechaIni, FechaFin", DbSystem, adOpenStatic
    If RsCrono.RecordCount = 0 Then
        Op1(1).Value = True
        Op1(0).Enabled = False
        DGCrono.Visible = False
    End If
    Set DGCrono.DataSource = RsCrono
    With DGCrono
        .Columns("CodCCosto").Visible = False
        .Columns("ID_FechaPago").Visible = False
        .Columns("Nombre").Width = Int(.Columns("Nombre").Width * 1.5)
        .Columns("Descripcion").Width = Int(.Columns("Descripcion").Width * 1.5)
    End With
    xTipo.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RsCrono = Nothing
End Sub

Private Sub op1_Click(Index As Integer)
    Select Case Index
        Case 0
            xCCosto.Locked = True
            xNombre.Locked = True
            xFechaIni.Enabled = False
            xFechaFin.Enabled = False
            DGCrono.Visible = True
        Case 1
            xCCosto.Locked = False
            xNombre.Locked = False
            xFechaIni.Enabled = True
            xFechaFin.Enabled = True
            DGCrono.Visible = False
    End Select
End Sub

Private Sub RsCrono_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    If adReason = adRsnRequery Then Exit Sub
    If RsCrono.RecordCount = 0 Or RsCrono.EOF Then Exit Sub
    xCCosto.Text = RsCrono!CodCCosto & ": " & RsCrono!Nombre
    xCCosto.Tag = RsCrono!CodCCosto
    xFechaIni.Value = RsCrono!FechaIni
    xFechaFin.Value = RsCrono!FechaFin
    xNombre.Text = RsCrono!Descripcion
End Sub

Private Sub xCCosto_DblClick()
    If xCCosto.Locked = True Then Exit Sub
    Dim RsCCostos As New ADODB.Recordset
    RsCCostos.Open "Select CodCCosto,Nombre From CCostos Order By CodCCosto", DbSystem, adOpenKeyset, adLockOptimistic
    frmComun.Conectar RsCCostos
    frmComun.Show 1
    If vgUtil(1) <> "" Then
        xCCosto.Text = vgUtil(1) & " :  " & vgUtil(2)
        xCCosto.Tag = vgUtil(1)
    End If
    Set RsCCostos = Nothing
End Sub
