VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAutoAd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorellenado de Adelantos"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frAutoAd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5295
   Begin VB.CommandButton CmdDetalle 
      Caption         =   "&Detallar"
      Height          =   345
      Left            =   885
      TabIndex        =   10
      Top             =   2775
      Width           =   1245
   End
   Begin VB.CommandButton CmdConcepto 
      Caption         =   "&Concepto"
      Height          =   405
      Left            =   315
      TabIndex        =   9
      Top             =   3375
      Visible         =   0   'False
      Width           =   1080
   End
   Begin MSScriptControlCtl.ScriptControl vbScript1 
      Left            =   120
      Top             =   2745
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   345
      Left            =   3975
      TabIndex        =   8
      Top             =   2745
      Width           =   1245
   End
   Begin VB.CommandButton cmProcesar 
      Caption         =   "&Procesar"
      Height          =   345
      Left            =   2670
      TabIndex        =   7
      Top             =   2745
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      Caption         =   "Autorellenado"
      Height          =   2490
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   5100
      Begin AplisetControlText.Aplitext xPorc 
         Height          =   285
         Left            =   2970
         TabIndex        =   2
         Top             =   383
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xFormula 
         Height          =   315
         Left            =   210
         TabIndex        =   6
         Top             =   1590
         Visible         =   0   'False
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         Text            =   ""
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Formula de cálculo (No incluye IF - IFF)"
         Height          =   210
         Left            =   195
         TabIndex        =   5
         Top             =   1215
         Width           =   3120
      End
      Begin VB.OptionButton Option2 
         Caption         =   "En nuevos soles (S/.)"
         Height          =   210
         Left            =   195
         TabIndex        =   3
         Top             =   817
         Width           =   1875
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Porcentaje del Basico"
         Height          =   210
         Left            =   195
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   1875
      End
      Begin AplisetControlText.Aplitext xSoles 
         Height          =   285
         Left            =   2970
         TabIndex        =   4
         Top             =   780
         Visible         =   0   'False
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
   End
   Begin MSDataGridLib.DataGrid DtgDetalle 
      Height          =   2055
      Left            =   1665
      TabIndex        =   11
      Top             =   3270
      Visible         =   0   'False
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   3625
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
End
Attribute VB_Name = "frAutoAd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MES_X As String
Public ANNO_X As String
Dim ANNO As String
Private Sub CMCANCELAR_CLICK()
    C = 0
    Unload Me
End Sub

Private Sub CmdConcepto_Click()
'Obterner los conceptos de ingresos y egresos para detallar el adelanto
Dim RSCONCEP As New ADODB.Recordset
    RSCONCEP.Open "SELECT CODIGO, NOMBRE,COMENTARIO,TIPO,FORMULA,COLPLANILLA FROM CONCEPTOS WHERE TIPO>0 AND TIPO<3 ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCONCEP
    C = 1
    frmComun.TOP = 1000
    frmComun.Left = 5000
    frmComun.Show 1
End Sub
Private Sub CmdDetalle_Click()
Me.Height = 5900
DtgDetalle.Visible = True
CmdConcepto.Visible = True
CmdConcepto.SetFocus
End Sub

Private Sub CMPROCESAR_Click()
Dim RSCAMP As New ADODB.Recordset
Dim ULTPOS
    If Option1.Value Then
        If Val(xPorc.Text) <= 0 Then
            MsgBox "ERROR: El porcentaje de cálculo no debe ser menor o igual a cero", vbCritical
            Exit Sub
        End If
        Dim RSADEL As New ADODB.Recordset
'solucion
    If VAR = 1 Then
        Set RSADEL = frAdelantos.GETDATA
    Else
        Set RSADEL = FrmAdeldet.GETDATA
    End If
        With RSADEL
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                !MONTO = !BASICO * Val(xPorc.Text) / 100
                .MoveNext
            Loop
            .Bookmark = ULTPOS
        End With
        Set RSADEL = Nothing
    End If
    If Option2.Value Then
        If Val(xSoles.Text) <= 0 Then
            MsgBox "ERROR: La cantidad en nuevos soles no debe ser menor o igual a cero", vbCritical
            Exit Sub
        End If
        Dim RSADEL2 As New ADODB.Recordset
        If VAR = 1 Then
            Set RSADEL2 = frAdelantos.GETDATA
        Else
            Set RSADEL2 = FrmAdeldet.GETDATA
        End If
        With RSADEL2
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                !MONTO = Val(xSoles.Text)
                .MoveNext
            Loop
            .Bookmark = ULTPOS
        End With
        Set RSADEL2 = Nothing
    End If
    If Option3.Value Then
        On Error GoTo ERRVBS
        If xFormula.Text = "" Then
            MsgBox "ERROR: Debera ingresar una formula aritmética", vbCritical
            Exit Sub
        End If
        Dim RSADEL3 As New ADODB.Recordset
        If VAR = 1 Then
            Set RSADEL3 = frAdelantos.GETDATA
        Else
            Set RSADEL3 = FrmAdeldet.GETDATA
        End If
        Dim VALOR
        With RSADEL3
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                vbScript1.AddCode "BASICO=" & !BASICO
                VALOR = vbScript1.Eval("" & xFormula.Text)
                !MONTO = Val(VALOR)
                .MoveNext
                vbScript1.Reset
            Loop
            .Bookmark = ULTPOS
        End With
        Set RSADEL3 = Nothing
    End If
    Unload Me
    Exit Sub
ERRVBS:
    MsgBox "La fórmula que ha ingresado es incorrecta", vbCritical
    Set RSADEL3 = Nothing
    Exit Sub
End Sub

Private Sub COMMAND1_CLICK()
    frmHelpTmp.LlamaFrm = 2
    frmHelpTmp.Show 1
End Sub

Private Sub Form_Load()
 If ExisteTablaAux("##TMPDETALLE") Then DBAUXCOM.Execute "DROP TABLE ##TMPDETALLE"
      DBAUXCOM.Execute "CREATE TABLE ##TMPDETALLE (NOMBRE VARCHAR(30), TIPO INT)"
  End Sub

Private Sub OPTION1_CLICK()
    xPorc.Visible = True
    xFormula.Visible = False
    xSoles.Visible = False
    xPorc.SetFocus
End Sub

Private Sub OPTION1_KeyDown(KEYCODE As Integer, SHIFT As Integer)
    If KEYCODE = 13 Then xPorc.SetFocus
End Sub

Private Sub OPTION2_Click()
    xPorc.Visible = False
    xFormula.Visible = False
    xSoles.Visible = True
    xSoles.SetFocus
End Sub

Private Sub OPTION3_Click()
    xPorc.Visible = False
    xFormula.Visible = True
    xSoles.Visible = False
    xFormula.SetFocus
End Sub
