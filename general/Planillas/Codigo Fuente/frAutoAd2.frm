VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAutoAd2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorellenado de Adelantos"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "frAutoAd2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5295
   Begin MSScriptControlCtl.ScriptControl vbScript1 
      Left            =   135
      Top             =   2670
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
      Left            =   2685
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
End
Attribute VB_Name = "frAutoAd2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public MES_X As String
Public ANNO_X As String
Dim ANNO As String

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CmdConcepto_Click()
Dim RSCONCEP As New ADODB.Recordset
    RSCONCEP.Open "SELECT CODIGO, NOMBRE,COMENTARIO,TIPO,FORMULA,COLPLANILLA FROM CONCEPTOS WHERE TIPO>0 AND TIPO<3 ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCONCEP
    frmComun.Show 1
    'If VGUTIL(1) <> "" Then
    '    AgregaCnpt.Tag = VGUTIL(1)
    'Else
    '    Set RSCONCEP = Nothing
    '    Exit Sub
    'End If
frConcpt.Show 1
End Sub

Private Sub CmdDetalle_Click()
Me.Height = 5900
'DtgDetalle.Visible = True
'CmdConcepto.Visible = True
'CmdConcepto.SetFocus
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
        Screen.MousePointer = 11
        Set RSADEL = FrmAdeldet.GETDATA
        With RSADEL
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                If FrmAdeldet.Columna = 0 Then Screen.MousePointer = 1: Exit Sub
                .Fields(FrmAdeldet.Columna) = .Fields(FrmAdeldet.Columna - 1) * Val(xPorc.Text) / 100
                .MoveNext
            Loop
            .Bookmark = ULTPOS
        End With
        Set RSADEL = Nothing
    End If
    If Option2.Value Then
        If Val(xSoles.Text) < 0 Then
            MsgBox "ERROR: La cantidad en nuevos soles no debe ser menor o igual a cero", vbCritical
            Exit Sub
        End If
        Dim RSADEL2 As New ADODB.Recordset
        Set RSADEL2 = FrmAdeldet.GETDATA
        With RSADEL2
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                .Fields(FrmAdeldet.Columna) = Val(xSoles.Text)
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
        Set RSADEL3 = FrmAdeldet.GETDATA
        Dim VALOR
        With RSADEL3
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                vbScript1.AddCode "BASICO=" & !BASICO
                VALOR = vbScript1.Eval("" & xFormula.Text)
                .Fields(FrmAdeldet.Columna) = Val(VALOR)
                .MoveNext
                vbScript1.Reset
            Loop
            .Bookmark = ULTPOS
        End With
        Set RSADEL3 = Nothing
    End If
    
    Call SumaAdel
    
    Unload Me
    Screen.MousePointer = 1
    Exit Sub
ERRVBS:
    MsgBox "La fórmula que ha ingresado es incorrecta", vbCritical
    Set RSADEL3 = Nothing
    Exit Sub
End Sub
Public Sub SumaAdel(Optional RS As ADODB.Recordset)
    'SUMAR LOS TOTALES DE ADELANTO
Dim RSADEL2 As ADODB.Recordset
Dim I As Integer
Dim ACUM As Double
Dim ULTPOS As Variant
Dim CONCEP As Double
Dim TIPO As String
    Set RSADEL2 = New ADODB.Recordset
    If RS Is Nothing Then
        Set RSADEL2 = FrmAdeldet.GETDATA
      Else
        Set RSADEL2 = RS
    End If
        With RSADEL2
            ULTPOS = .Bookmark
            .MoveFirst
            Do While Not .EOF
                ACUM = 0
                For I = 0 To RSADEL2.Fields.Count - 1
                    If I < 5 Then GoTo FINAL
                    If Not (.Fields(I).Name = "TOTAL ADELANTO" Or Left(.Fields(I).Name, 2) = "M ") Then
                        CONCEP = ESNULO(.Fields(I).Value, 0)
                        If .Fields(I).Name <> "MONTO" Then
                            TIPO = DevuelveValor("SELECT TIPO FROM CONFIADEL WHERE NOMBRE='" & .Fields(I).Name & "'", DBSYSTEM)
                            Select Case TIPO
                                Case 2: CONCEP = CONCEP * -1
                            End Select
                        End If
                        ACUM = ACUM + CONCEP
                    End If
FINAL:
                Next
                 .Fields("TOTAL ADELANTO") = ACUM
                 FrmAdeldet.Refresh
                 .MoveNext
            Loop
            .Bookmark = ULTPOS
        End With
    Set RSADEL2 = Nothing
End Sub

Private Sub Command1_Click()
    frmHelpTmp.LlamaFrm = 2
    frmHelpTmp.Show 1
End Sub

Private Sub OPTION1_CLICK()
    xPorc.Visible = True
    xFormula.Visible = False
    xSoles.Visible = False
    xPorc.SetFocus
End Sub

Private Sub OPTION1_KeyDown(KEYCODE As Integer, Shift As Integer)
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

