VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmdebPend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debitos de Cta. Cte. Pendientes"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3645
   Icon            =   "FrmdebPend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   2250
      TabIndex        =   6
      Text            =   "SqlCad"
      Top             =   1200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   2265
      TabIndex        =   5
      Top             =   855
      Width           =   1140
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2295
      TabIndex        =   4
      Top             =   315
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1335
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   1935
      Begin Crystal.CrystalReport Reporte 
         Left            =   1440
         Top             =   270
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.OptionButton xTodos 
         Caption         =   "Todos"
         Height          =   300
         Left            =   345
         TabIndex        =   3
         Top             =   915
         Width           =   1050
      End
      Begin VB.OptionButton xEgresos 
         Caption         =   "&Egresos"
         Height          =   210
         Left            =   345
         TabIndex        =   2
         Top             =   645
         Width           =   1050
      End
      Begin VB.OptionButton XIngresos 
         Caption         =   "&Ingresos"
         Height          =   240
         Left            =   345
         TabIndex        =   1
         Top             =   330
         Width           =   1050
      End
   End
End
Attribute VB_Name = "FrmdebPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDACEPTAR_CLICK()
    Dim X As Long
    Dim OPC1 As String
    Dim TIPO As String
    Screen.MousePointer = 11
    OPC1 = ""
    TIPO = "TODOS"
    If XIngresos.Value Then
        OPC1 = " AND MOVICTA.TIPOGRUPO=1"
        TIPO = UCase("INGRESOS")
    End If
    If xEgresos.Value Then
        OPC1 = " AND MOVICTA.TIPOGRUPO=2"
        TIPO = UCase("EGRESOS")
    End If
    If xTodos.Value Then
        TIPO = UCase("TODOS")
        OPC1 = ""
    End If
        
    Dim RUTA As String
    
    If ExisteTablaAux(" [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] "
    SqlCad.Text = "" & _
        "SELECT MOVICTA.CODGRUPO,CTAGRUPO.NOMBRE ," & _
        "SUM(MOVICTA.CAPITAL) AS CAPITAL, SUM(MOVICTA.SALDO)AS SALDO," & _
        "SUM(MOVICTA.CUOTA) AS CUOTA " & _
        " INTO   [##TMPGRUPCTEPEND" & VGL_COMPUTER & "]  " & _
        " FROM MOVICTA, CTAGRUPO " & _
        " WHERE " & _
        "    MOVICTA.CODGRUPO=CTAGRUPO.CODGRUPO AND " & _
        "    MOVICTA.SALDO<>0 " & OPC1 & _
        " GROUP BY MOVICTA.CODGRUPO,CTAGRUPO.NOMBRE"
    DBSYSTEM.Execute SqlCad.Text, X
    Screen.MousePointer = 1
    If X = 0 Then
      MsgBox "MENSAJE DEL SISTEMA: " & _
      " NO SE ENCONTRARÛN REGISTROS ", vbInformation
      Exit Sub
    End If
    With Reporte
        .ReportFileName = REGSISTEMA.REPORTES & "PLAN0015.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = " [##TMPGRUPCTEPEND" & VGL_COMPUTER & "] "
        .Destination = crptToWindow
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "PLAN0015 - CONSOLIDADO DE SALDOS PENDIENTES POR GRUPO DE CTA."
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XTIPO='" & TIPO & "'"
        .Formulas(2) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
End Sub
Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub


