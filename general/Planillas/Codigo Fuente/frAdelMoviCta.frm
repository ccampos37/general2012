VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frAdelMoviCta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargos de Cuenta Corriente a Adelantos"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "frAdelMoviCta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   4950
      TabIndex        =   5
      ToolTipText     =   "Sólo para eliminar Ctas. Ctes. de Cuotas Fijas."
      Top             =   5190
      Width           =   1230
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3562
      TabIndex        =   2
      Top             =   5190
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2175
      TabIndex        =   1
      Top             =   5190
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid DGMovis 
      Height          =   4080
      Left            =   150
      TabIndex        =   0
      Top             =   975
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   7197
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
      Caption         =   "Cuentas Corrientes de Trabajadores"
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
   Begin VB.Label Label2 
      Caption         =   $"frAdelMoviCta.frx":030A
      Height          =   660
      Left            =   165
      TabIndex        =   6
      Top             =   255
      Width           =   4605
   End
   Begin VB.Label Label3 
      Caption         =   "Tip  I = Ingreso Tip E = Egreso"
      Height          =   405
      Left            =   150
      TabIndex        =   4
      Top             =   5145
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Adelantos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   4920
      TabIndex        =   3
      Top             =   585
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5505
      Picture         =   "frAdelMoviCta.frx":03B6
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frAdelMoviCta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FORMULARIO LLAMADO DESDE FRADELANTOS
Option Explicit
Dim RSMOVIS As New ADODB.Recordset
Private Sub Command1_Click()
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    RSMOVIS.Save App.PATH & "\ADELCC.DYB", adPersistADTG
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub COMMAND4_Click()
'    If MsgBox("DESEA QUITAR TODOS LOS CARGOS DE CUENTAS CORRIENTE AL GRUPO DE ADELANTOS DE REMUNERACIONES QUE ESTA EDITANDO", vbYesNo + vbQuestion) = vbNo Then Exit Sub
'    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then Kill App.PATH & "\ADELCC.DYB"
    If ESNULO(RSMOVIS!SECUENCIA, 0) = 0 Then
        If MsgBox("Esta seguro que desea quitar el debito", vbQuestion + vbYesNo) = vbYes Then
            RSMOVIS.Delete adAffectCurrent
        End If
      Else
        MsgBox "Los debitos programados no se pueden eliminar", vbExclamation
    End If
End Sub

Private Sub DGMOVIS_GOTFOCUS()
    'IF RSMOVIS.RECORDCOUNT > 0 THEN DGMOVIS.COL = DGMOVIS.COLUMNS("DEBITO").COLINDEX
End Sub

Private Sub DGMOVIS_HEADCLICK(ByVal COLINDEX As Integer)
    RSMOVIS.Sort = DGMovis.Columns(COLINDEX).Caption
End Sub

Private Sub Form_Load()
    Dim CAD As String
On Error GoTo handler
    'If ExisteTablaAux("[##TMPDEB" & VGL_COMPUTER & "]") Then DBSYSTEM.Execute "DROP TABLE [##TMPDEB" & VGL_COMPUTER & "]"
    CAD = "SELECT * FROM (SELECT " & _
        "CODMOV,VWTRABAJ.CODTRAB,NOMBRES, " & _
        "TIP = CASE TIPOGRUPO WHEN 1 THEN 'I' ELSE 'E' END, " & _
        "CAPITAL, " & _
        "DEBITO=(CUOTA*(PORCQUINC/100)),DESCRIPCION,SECUENCIA=MOVICTA.ULTSECU " & _
        "From MOVICTA, VWTRABAJ " & _
        "Where " & _
        "MOVICTA.CODTRAB=VWTRABAJ.CODTRAB AND PORCQUINC<>0 AND " & _
        "FECHAINI <=" & FechS(REGINPUT.FECHAFIN, Sqlf) & " AND SALDO >= 0 AND " & _
        "MOVICTA.PROGRAMADO = 0 AND " & _
        "VWTRABAJ.CODTRAB IN (SELECT CODTRAB FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] ) " & _
        "Union " & _
        "SELECT " & _
           " MOVICTA.CODMOV,VWTRABAJ.CODTRAB,NOMBRES, " & _
           " TIP = CASE TIPOGRUPO WHEN 1 THEN 'I' ELSE 'E' END, " & _
           " CAPITAL, " & _
           " DEBITO=CTACTEPROG.IMPORTE,DESCRIPCION,SECUENCIA=CTACTEPROG.SECUENCIA " & _
           " From MOVICTA, VWTRABAJ, CTACTEPROG " & _
           " Where " & _
           " MOVICTA.CODTRAB=VWTRABAJ.CODTRAB  AND " & _
           " CTACTEPROG.FECHA BETWEEN  " & FechS(REGINPUT.FECHAINI, Sqlf) & " AND " & FechS(REGINPUT.FECHAFIN, Sqlf) & "  AND " & _
           " CTACTEPROG.CODMOV=MOVICTA.CODMOV AND " & _
           " MITAPER='V' AND SALDO >= 0 AND " & _
           " MOVICTA.PROGRAMADO =1  AND " & _
           " VWTRABAJ.CODTRAB IN (SELECT CODTRAB FROM  [##_TMPADELANTO" & VGL_COMPUTER & "] )) AS XX WHERE XX.CODMOV NOT IN (SELECT CODMOV FROM PAGOSCTA WHERE TIPOBOLETA='A' AND CODNOMBOL=" & FrmAdeldet.Lista.SelectedItem.Tag & ")"
    Dim XTABLA As String
    XTABLA = "[TMP" & REGSISTEMA.USER & "ADEL]"
    If UCase(Dir$(App.PATH & "\ADELCC.DYB")) = "ADELCC.DYB" Then
        RSMOVIS.Open App.PATH & "\ADELCC.DYB", , adOpenKeyset, adLockOptimistic, adCmdFile
    Else
        RSMOVIS.Open CAD, DBSYSTEM, adOpenStatic, adLockOptimistic
    End If
    Set RSMOVIS.ActiveConnection = Nothing
    REFRESCAR
    If RSMOVIS.RecordCount = 0 Then
        Command1.Enabled = False
        Command4.Enabled = False
    End If

Exit Sub
handler:
Exit Sub
Resume
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSMOVIS = Nothing
End Sub

Private Sub IMAGE1_CLICK()
    MsgBox "DESARROLLADO POR FERNANDO COSSIO : FONO 548 - 4925", vbInformation
End Sub

Public Sub REFRESCAR()
    Set DGMovis.DataSource = RSMOVIS
    With DGMovis
        .Columns("CODTRAB").Visible = False
        .Columns("CODMOV").Visible = False
        .Columns("TIP").Alignment = dbgCenter
        .Columns("DEBITO").Alignment = dbgRight
        .Columns("DEBITO").NumberFormat = "0.00 "
        .Columns("CAPITAL").Alignment = dbgRight
        .Columns("CAPITAL").NumberFormat = "0.00 "
        .Columns("DESCRIPCION").Width = 2819.906
        .Columns("DEBITO").Width = 975.1182
        .Columns("CAPITAL").Width = 1065.26
        .Columns("NOMBRES").Width = 2369.764
        .Columns("TIP").Width = 329.9528
        .Columns("TIP").Locked = True
        .Columns("NOMBRES").Locked = True
        .Columns("CAPITAL").Locked = True
        .Columns("DESCRIPCION").Locked = True
        .Refresh
    End With
End Sub

