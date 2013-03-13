VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frTrb5ta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trabajadores afectos a 5ta. Categoria"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "frTrb5ta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmAgregar 
      Caption         =   "&Agregar trabajador"
      Height          =   345
      Left            =   105
      TabIndex        =   7
      Top             =   6105
      Width           =   1890
   End
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   5505
      TabIndex        =   8
      Top             =   6105
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trabajador"
      Height          =   2070
      Left            =   105
      TabIndex        =   1
      Top             =   3945
      Width           =   6645
      Begin VB.CommandButton cmActualizar 
         Caption         =   "Ac&tualizar"
         Height          =   330
         Left            =   150
         TabIndex        =   11
         Top             =   1680
         Width           =   1305
      End
      Begin AplisetControlText.Aplitext xRemuPrevia 
         Height          =   300
         Left            =   5505
         TabIndex        =   26
         Top             =   1635
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xRentaPrevia 
         Height          =   300
         Left            =   5505
         TabIndex        =   27
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xRemu3ros 
         Height          =   300
         Left            =   5505
         TabIndex        =   28
         Top             =   1005
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xGN 
         Height          =   300
         Left            =   1830
         TabIndex        =   29
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xGF 
         Height          =   300
         Left            =   1830
         TabIndex        =   30
         Top             =   1005
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xTrab 
         Height          =   300
         Left            =   1200
         TabIndex        =   31
         Top             =   300
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Remuneración Acumulada Previa"
         Height          =   195
         Index           =   6
         Left            =   3015
         TabIndex        =   12
         Top             =   1695
         Width           =   2370
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Renta Acumulada Previa"
         Height          =   195
         Index           =   5
         Left            =   3015
         TabIndex        =   15
         Top             =   1388
         Width           =   1770
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Remuneración por Terceros"
         Height          =   195
         Index           =   4
         Left            =   3015
         TabIndex        =   16
         Top             =   1073
         Width           =   1980
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Grat. Navidad"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   1388
         Width           =   990
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Grat.Fiestas Patrias"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   1073
         Width           =   1365
      End
      Begin VB.Label xDatos 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   1200
         TabIndex        =   4
         Top             =   615
         Width           =   4230
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Datos"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Top             =   645
         Width           =   420
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   345
         Width           =   765
      End
   End
   Begin MSDataGridLib.DataGrid DGTrabs 
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   3720
      _ExtentX        =   6562
      _ExtentY        =   6668
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
      Caption         =   "Trabajadores afectos"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   0
      Left            =   5565
      TabIndex        =   32
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   0
      Left            =   4485
      TabIndex        =   33
      Top             =   315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   1
      Left            =   4485
      TabIndex        =   34
      Top             =   615
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   2
      Left            =   4485
      TabIndex        =   35
      Top             =   915
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   3
      Left            =   4485
      TabIndex        =   36
      Top             =   1215
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   4
      Left            =   4485
      TabIndex        =   37
      Top             =   1515
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   5
      Left            =   4485
      TabIndex        =   38
      Top             =   1815
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   6
      Left            =   4485
      TabIndex        =   39
      Top             =   2115
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   7
      Left            =   4485
      TabIndex        =   40
      Top             =   2415
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   8
      Left            =   4485
      TabIndex        =   41
      Top             =   2715
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   9
      Left            =   4485
      TabIndex        =   42
      Top             =   3015
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   10
      Left            =   4485
      TabIndex        =   43
      Top             =   3315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext Aplitext2 
      Height          =   285
      Index           =   11
      Left            =   4485
      TabIndex        =   44
      Top             =   3615
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   1
      Left            =   5565
      TabIndex        =   45
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   615
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   2
      Left            =   5565
      TabIndex        =   46
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   915
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   3
      Left            =   5565
      TabIndex        =   47
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   1215
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   4
      Left            =   5565
      TabIndex        =   48
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   1515
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   5
      Left            =   5565
      TabIndex        =   49
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   1815
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   6
      Left            =   5565
      TabIndex        =   50
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   2115
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   7
      Left            =   5565
      TabIndex        =   51
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   2415
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   8
      Left            =   5565
      TabIndex        =   52
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   2715
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   9
      Left            =   5565
      TabIndex        =   53
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   3015
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   10
      Left            =   5565
      TabIndex        =   54
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   3315
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin AplisetControlText.Aplitext RAcu 
      Height          =   285
      Index           =   11
      Left            =   5565
      TabIndex        =   55
      ToolTipText     =   "Renta de 5ta. Categoria"
      Top             =   3615
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   503
      Text            =   "0"
      TipoDato        =   "N"
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mes     Remuneración       Renta"
      Height          =   240
      Left            =   3960
      TabIndex        =   23
      Top             =   60
      Width           =   2670
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "DIC"
      Height          =   285
      Index           =   11
      Left            =   3960
      TabIndex        =   24
      Top             =   3615
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NOV"
      Height          =   285
      Index           =   10
      Left            =   3960
      TabIndex        =   25
      Top             =   3315
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "OCT"
      Height          =   285
      Index           =   9
      Left            =   3960
      TabIndex        =   22
      Top             =   3015
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SET"
      Height          =   285
      Index           =   8
      Left            =   3960
      TabIndex        =   21
      Top             =   2715
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGO"
      Height          =   285
      Index           =   7
      Left            =   3960
      TabIndex        =   18
      Top             =   2415
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUL"
      Height          =   285
      Index           =   6
      Left            =   3960
      TabIndex        =   17
      Top             =   2115
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "JUN"
      Height          =   285
      Index           =   5
      Left            =   3960
      TabIndex        =   14
      Top             =   1815
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MAY"
      Height          =   285
      Index           =   4
      Left            =   3960
      TabIndex        =   13
      Top             =   1515
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ABR"
      Height          =   285
      Index           =   3
      Left            =   3960
      TabIndex        =   10
      Top             =   1215
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MAR"
      Height          =   285
      Index           =   2
      Left            =   3960
      TabIndex        =   9
      Top             =   915
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FEB"
      Height          =   285
      Index           =   1
      Left            =   3960
      TabIndex        =   6
      Top             =   615
      Width           =   510
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ENE"
      Height          =   285
      Index           =   0
      Left            =   3960
      TabIndex        =   5
      Top             =   315
      Width           =   510
   End
End
Attribute VB_Name = "frTrb5ta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents RSTRAB As ADODB.Recordset
Attribute RSTRAB.VB_VarHelpID = -1

Private Sub CMACTUALIZAR_CLICK()
    Dim X As Integer
    Dim STR1 As String
    STR1 = ""
    If cmAgregar.Visible And RSTRAB.EOF Then Exit Sub
    If cmAgregar.Visible Then 'QUIERE DECIR QUE NO AGREGABA UNO, PUES CUANDO SE AGREGA EL BOTON AGREGAR DESAPARECE
        For X = 0 To 11
            STR1 = STR1 + "R" & (X + 1) & "=" & Aplitext2(X).Text & ", RA" & (X + 1) & "=" & RAcu(X).Text & ","
        Next
        DBSYSTEM.Execute "UPDATE RETEN5TA SET GF=" & xGF.Text & ", GN=" & xGN.Text & "," & STR1 & "REMU3ROS=" & xRemu3ros.Text & ", RENTAPREVIA=" & xRentaPrevia.Text & ", REMUPREVIA=" & xRemuPrevia.Text & " WHERE CODTRAB='" & RSTRAB!CODTRAB & "'"
    Else
        If xTrab.Tag = "" Then
            MsgBox "FALTA AGREGAR AL FORMULARIO EL TRABAJADOR QUE VA A ESTAR AFECTO A IMPUESTO A LA RENTA DE QUINTA CATEGORIA", vbCritical
            Exit Sub
        End If
        If Val(xGF.Text) <= 0 Then
            MsgBox "MONTO DE GRATIFICACIÓN ORDINARIA DE FIESTAS PATRIAS INVÁLIDO", vbCritical
            Exit Sub
        End If
        If Val(xGN.Text) <= 0 Then
            MsgBox "MONTO DE GRATIFICACIÓN ORDINARIA DE NAVIDAD INVÁLIDO", vbCritical
            Exit Sub
        End If
        For X = 0 To 11
            STR1 = STR1 + "," & Aplitext2(X).Text
        Next
        For X = 0 To 11
            STR1 = STR1 + "," & RAcu(X).Text
        Next
        DBSYSTEM.Execute "INSERT INTO RETEN5TA VALUES ('" & xTrab.Tag & "'," & DateSQL(Date) & "," & xGF.Text & "," & xGN.Text & STR1 & "," & xRemu3ros.Text & "," & xRentaPrevia.Text & "," & xRemuPrevia.Text & ")"
        cmAgregar.Visible = True
        DGTrabs.Visible = True
        RSTRAB.Requery
        Set DGTrabs.DataSource = RSTRAB
        frTrb5ta.DGTrabs.Columns("NOMBRES").Width = 2550.047
    End If
    MsgBox "LA ACTUALIZACIÓN SE HA COMPLETADO SATISFACTORIAMENTE", vbInformation
End Sub

Private Sub CMAGREGAR_CLICK()
    Dim X As Integer
    Dim STRVAL As String
    cmAgregar.Visible = False
    DGTrabs.Visible = False
    For X = 0 To 11
        Aplitext2(X).Text = 0
        RAcu(X).Text = 0
    Next
    xTrab.Tag = ""
    xTrab.Text = ""
    xDatos.Caption = ""
    xGF.Text = 0
    xGN.Text = 0
    xRemu3ros.Text = 0
    xRemuPrevia.Text = 0
    xRentaPrevia.Text = 0
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set RSTRAB = New ADODB.Recordset
    RSTRAB.Open "SELECT A.CODTRAB, B.NOMBRES FROM RETEN5TA A, VWTRABAJ B WHERE A.CODTRAB=B.CODTRAB ORDER BY NOMBRES", DBSYSTEM, adOpenStatic
    Set DGTrabs.DataSource = RSTRAB
    frTrb5ta.DGTrabs.Columns("NOMBRES").Width = 2550.047
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSTRAB = Nothing
End Sub

Private Sub RSTRAB_MOVECOMPLETE(ByVal ADREASON As ADODB.EventReasonEnum, ByVal PERROR As ADODB.Error, ADSTATUS As ADODB.EventStatusEnum, ByVal PRECORDSET As ADODB.Recordset)
    If RSTRAB.EOF Then Exit Sub
    If ADREASON = adRsnFirstChange Then Exit Sub
    Dim RS1 As New ADODB.Recordset
    RS1.Open "SELECT * FROM RETEN5TA WHERE CODTRAB='" & RSTRAB!CODTRAB & "'", DBSYSTEM, adOpenStatic
    With RS1
        xGF.Text = !GF
        xGN.Text = !GN
        Aplitext2(0).Text = !R1
        Aplitext2(1).Text = !R2
        Aplitext2(2).Text = !R3
        Aplitext2(3).Text = !R4
        Aplitext2(4).Text = !R5
        Aplitext2(5).Text = !R6
        Aplitext2(6).Text = !R7
        Aplitext2(7).Text = !R8
        Aplitext2(8).Text = !R9
        Aplitext2(9).Text = !R10
        Aplitext2(10).Text = !R11
        Aplitext2(11).Text = !R12
        RAcu(0).Text = !RA1
        RAcu(1).Text = !RA2
        RAcu(2).Text = !RA3
        RAcu(3).Text = !RA4
        RAcu(4).Text = !RA5
        RAcu(5).Text = !RA6
        RAcu(6).Text = !RA7
        RAcu(7).Text = !RA8
        RAcu(8).Text = !RA9
        RAcu(9).Text = !RA10
        RAcu(10).Text = !RA11
        RAcu(11).Text = !RA12
        xRemu3ros.Text = !REMU3ROS
        xRemuPrevia.Text = !REMUPREVIA
        xRentaPrevia.Text = !RENTAPREVIA
        xTrab.Text = RSTRAB!NOMBRES
    End With
    Set RS1 = Nothing
End Sub

Private Sub XTRAB_DBLCLICK()
    If cmAgregar.Visible Then
        MsgBox "NO SE PUEDE SELECCIONAR DESDE EL MÉTODO DE EDICIÓN", vbCritical
        Exit Sub
    End If
    Dim RS1 As New ADODB.Recordset
    Dim X As Integer
    RS1.Open "VWTRABAJ", DBSYSTEM, adOpenStatic
    frmComun.CONECTAR RS1
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        DBSYSTEM.Execute "UPDATE RETEN5TA SET R12=R12 WHERE CODTRAB='" & VGUTIL(1) & "'", X
        If X <> 0 Then
            MsgBox "EL TRABAJADOR QUE USTED HA SELECCIONADO YA PRESENTA UN REGISTRO DE RETENCIONES DEL IMPUESTO A LA RENTA DE 5TA. CATEGORIA", vbCritical
            Set RS1 = Nothing
            Exit Sub
        End If
        xTrab.Text = VGUTIL(1) & " :  " & VGUTIL(2)
        xTrab.Tag = VGUTIL(1)
        xDatos.Caption = RS1!CENTRO
    End If
    Set RS1 = Nothing
End Sub

