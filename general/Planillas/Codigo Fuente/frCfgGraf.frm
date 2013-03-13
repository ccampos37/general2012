VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frCfgGraf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del Gráfico"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frCfgGraf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Reporte 
      Left            =   2460
      Top             =   3690
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2325
      TabIndex        =   5
      Top             =   4605
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   4485
      Begin VB.Label Label1 
         Caption         =   "Aqui se muestran  las columnas de planilla que va porder comparar según los periodos seleccionados en el formulario anterior"
         Height          =   585
         Left            =   975
         TabIndex        =   6
         Top             =   165
         Width           =   3240
      End
      Begin VB.Image Image2 
         Height          =   330
         Left            =   120
         Picture         =   "frCfgGraf.frx":08CA
         Stretch         =   -1  'True
         Top             =   195
         Width           =   375
      End
   End
   Begin VB.TextBox SQL 
      Height          =   285
      Left            =   1125
      TabIndex        =   3
      Text            =   "SQL"
      Top             =   3105
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox SqlCad 
      Height          =   285
      Left            =   90
      TabIndex        =   2
      Text            =   "SqlCad"
      Top             =   3105
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2940
      Top             =   3570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCfgGraf.frx":1194
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   1005
      TabIndex        =   1
      Top             =   4605
      Width           =   1215
   End
   Begin MSComctlLib.ListView LColumnas 
      Height          =   3570
      Left            =   45
      TabIndex        =   0
      Top             =   945
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6297
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Suma"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4050
      Picture         =   "frCfgGraf.frx":14E8
      Top             =   4560
      Width           =   480
   End
End
Attribute VB_Name = "frCfgGraf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDACEPTAR_CLICK()
    Dim FLAG
    Dim CONCEPTO As String
    Screen.MousePointer = 11
    Call GRAFICA(FLAG)
    If Not FLAG Then Exit Sub
    Dim RSAUX As New ADODB.Recordset
    Dim RSGRAF As New ADODB.Recordset
    Dim I As Integer
    RSAUX.Open " [##TMPPLANGROUP" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    RSAUX.MoveFirst
    
    If ExisteTablaAux(" [##TMPGRAPLAN" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPGRAPLAN" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "" & _
    "CREATE TABLE  [##TMPGRAPLAN" & VGL_COMPUTER & "]  ( CODIGO VARCHAR(20), DESCRI VARCHAR(30),PERIODO VARCHAR(20), CANT  Numeric(20,2) , FECHA DATETIME )"
        
    RSGRAF.Open " [##TMPGRAPLAN" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Do While Not RSAUX.EOF
        For I = 1 To RSAUX.Fields.Count - 1
            RSGRAF.AddNew
            RSGRAF!PERIODO = Format(RSAUX!MES, "MMM - YYYY")
            RSGRAF!FECHA = FechS(RSAUX!MES, Adof)
            RSGRAF!Codigo = RSAUX.Fields(I).Name
            CONCEPTO = DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & Mid(RSAUX.Fields(I).Name, 2, 100) & "'", DBSYSTEM)
            If CONCEPTO = "" Then
                RSGRAF!DESCRI = Mid(RSAUX.Fields(I).Name, 2, 100)
              Else:
                RSGRAF!DESCRI = CONCEPTO
            End If
            RSGRAF!CANT = RSAUX.Fields(I).Value
            RSGRAF.Update
        Next
        RSAUX.MoveNext
    Loop
    
    With Reporte
        frWait.Show 1
        .Reset
        .WindowTitle = "PLAN0072.RPT - GRÁFICA ESTADISTICA"
        .ReportFileName = REGSISTEMA.REPORTES & "\PLAN0072.RPT"
        .Connect = "DSN=" & VGL_SERVERREP & ";User=sa;PWD=;DSQ=planilla_pp"
        .StoredProcParam(0) = VGL_COMPUTER
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowShowPrintBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowShowPrintSetupBtn = True
        .Formulas(0) = "XEMPRESA='" & REGSISTEMA.EMPRESA & "'"
        .Formulas(1) = "XRUC='" & REGSISTEMA.RUC & "'"
        If .Status <> 2 Then .Action = 1
    End With
    Screen.MousePointer = 1
End Sub

Private Sub CMDCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub Form_Load()
Dim RSAUX As New ADODB.Recordset
Dim CONCEPTO As String
Dim I As Integer
   Dim XLIST As ListItem
    RSAUX.Open Trim(REGSISTEMA.TABLAPLAN), DBSYSTEM
    RSAUX.MoveFirst
    For I = 0 To RSAUX.Fields.Count - 1
        Select Case RSAUX.Fields(I).Type
            Case 14, 5, 131, 4, 6
                Set XLIST = LColumnas.ListItems.Add(, , Trim(RSAUX.Fields(I).Name), , 1)
                CONCEPTO = DevuelveValor("SELECT NOMBRE FROM COLUMPL WHERE CODIGO='" & Trim(RSAUX.Fields(I).Name) & "'", DBSYSTEM)
                If CONCEPTO = "" Then
                    XLIST.SubItems(1) = Trim(RSAUX.Fields(I).Name)
                  Else:
                    XLIST.SubItems(1) = CONCEPTO
                End If
                XLIST.SubItems(2) = "SUM(" & RSAUX.Fields(I).Name & ") AS S" & Trim(RSAUX.Fields(I).Name)
        End Select
    Next
End Sub

Private Sub GRAFICA(Optional ByRef FLAG)
Dim I As Integer
Dim RUTA As String
    'VALIDANDO QUE AL MENOS ESTE MARCADO UN CAMPO
    For I = 1 To LColumnas.ListItems.Count
        If LColumnas.ListItems.Item(I).Checked Then
            FLAG = True
            Exit For
        End If
    Next

    If Not FLAG Then
        MsgBox "POR LO MENOS DEBE EXISTIR UNA COLUMNA PARA EJECUTAR EL REPORTE", vbInformation
        Screen.MousePointer = 1
        Exit Sub
    End If

    Screen.MousePointer = 11
    SQL.Text = ""
    'RUTA = " IN '" & REGSISTEMA.PATHEMPRESA & "\PLANILLA.MDB'  "
    SqlCad.Text = " WHERE MES IN ("
    'ARMANDO LA SELECCION DE LA PLANILLAS PARA LA GRAFICA
    With frPlans
        For I = 1 To .LPlans.ListItems.Count
            If .LPlans.ListItems.Item(I).Checked = True Then
                 SqlCad.Text = SqlCad.Text + _
                 DateSQL(FechaMMAAAA(Right(.LPlans.ListItems.Item(I).KEY, 6))) + ","
            End If
        Next
    End With
    SqlCad.Text = Left(SqlCad.Text, Len(SqlCad.Text) - 1)
    SqlCad.Text = SqlCad.Text + ")"
    
    'ARMAR CONSULTA PARA LA GRAFICA DE LOS CAMPOS SELECCIONADOS
    
    SQL.Text = "SELECT MES, "
    For I = 1 To LColumnas.ListItems.Count
        If LColumnas.ListItems.Item(I).Checked Then
            SQL.Text = SQL.Text + LColumnas.ListItems.Item(I).SubItems(2) + ","
        End If
    Next
    SQL.Text = Left(SQL.Text, Len(SQL.Text) - 1)
    SQL.Text = SQL.Text & " INTO  [##TMPPLANGROUP" & VGL_COMPUTER & "]  FROM " & REGSISTEMA.TABLAPLAN & RUTA & _
               SqlCad.Text & " GROUP BY MES "
    
    Dim X As Integer
    
    If ExisteTablaAux(" [##TMPPLANGROUP" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##TMPPLANGROUP" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute SQL.Text, X
    
    If X = 0 Then
        MsgBox "NO HAY REGISTROS PARA LA GRAFICA", vbExclamation
        Screen.MousePointer = 1
        Exit Sub
    End If
    
    Screen.MousePointer = 1
End Sub


