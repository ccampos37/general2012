VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrDaTr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de Datos"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "Seleccionar Todos"
      Height          =   300
      Left            =   45
      TabIndex        =   6
      Top             =   4305
      Width           =   2145
   End
   Begin VB.CommandButton cmSelecc 
      Caption         =   "Terminar"
      Height          =   360
      Left            =   3570
      TabIndex        =   2
      Top             =   3915
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<< Atras"
      Default         =   -1  'True
      Height          =   360
      Left            =   2370
      TabIndex        =   5
      Top             =   3915
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Eliminar los Actuales"
      Height          =   195
      Left            =   45
      TabIndex        =   3
      Top             =   4680
      Width           =   2100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   2730
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
            Picture         =   "FrDaTr.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LEmpresas 
      Height          =   2955
      Left            =   0
      TabIndex        =   0
      Top             =   870
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5212
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label xEmpresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   660
      TabIndex        =   7
      Top             =   465
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   615
      TabIndex        =   4
      Top             =   345
      Width           =   4125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccion de Datos a Trasladar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   660
      TabIndex        =   1
      Top             =   210
      Width           =   2700
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "FrDaTr.frx":015C
      Top             =   195
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   4830
   End
End
Attribute VB_Name = "FrDaTr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMSELECC_CLICK()
    'On Error Resume Next
    Dim RS As New ADODB.Recordset, RSCOMPARA As New ADODB.Recordset
    Dim CNX_AUX As ADODB.Connection
    Dim X As Long
    Dim TABLAIDEN As String
    Dim APELLIDOPAT, APELLIDOMAT, NOMBRES, TIPODOC, NDOC As String
    Dim FLAG As Integer
    Dim FILATXT As String
    Dim TABLA As String
    Screen.MousePointer = 11
    If Check1.Value = 1 Then
        If MsgBox("Ud primero eliminará los registros actuales. desea continuar?", vbYesNo, "Advertencia") = vbYes Then
            
            SQL = "DELETE FROM " & ModPlan.TRAS.TABLA
            DBSYSTEM.Execute SQL
            MsgBox "AHORA SE REALIZARÁ LA ACTUALIZACION", vbInformation
            TABLA = ModPlan.TRAS.TABLA
            For X = 1 To Me.LEmpresas.ListItems.Count
                If LEmpresas.ListItems(X).Checked = True Then
                    If ModPlan.TRAS.ESCADENA Then
                        SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT * FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                ModPlan.TRAS.TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "='" & LEmpresas.ListItems(X).Text & "'"
                    Else
                        If IDENTITY Then
                            If ExisteCampo("TIPO", ModPlan.TRAS.TABLA, DBSYSTEM) Then
                                ModPlan.TRAS.TABLA = ModPlan.TRAS.TABLA & "(NOMBRE, FORMULA, CRITERIO, TIPO)  "
                                SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT NOMBRE, FORMULA, CRITERIO, TIPO FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                            Else
                                ModPlan.TRAS.TABLA = ModPlan.TRAS.TABLA & "(NOMBRE, FORMULA, CRITERIO) "
                                SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT NOMBRE, FORMULA, CRITERIO FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                            End If
                        Else
                                SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT * FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                        End If
                    End If
                    DBSYSTEM.Execute SQLEXEC
                End If
            Next
        End If
    Else
        Set CNX_AUX = New ADODB.Connection
       CNX_AUX.CursorLocation = adUseClient
       CNX_AUX.CommandTimeout = 0
       CNX_AUX.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=SOPORTE;Password=SOPORTE;Initial Catalog=" & ModPlan.TRAS.RUTABASE & ";Data Source=" & VGL_SERVER
       CNX_AUX.Open
        'SET RSCOMPARA = NEW ADODB.RECORDSET
        'SET RS = NEW ADODB.RECORDSET
        For X = 1 To Me.LEmpresas.ListItems.Count
            If LEmpresas.ListItems(X).Checked = True Then
                If TRAS.ESCADENA Then
                    SQLAUX = "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE " & ModPlan.TRAS.FIELDK & _
                         "='" & LEmpresas.ListItems(X).Text & "'"
                Else
                    SQLAUX = "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE " & ModPlan.TRAS.FIELDK & _
                         "=" & LEmpresas.ListItems(X).Text
                End If
                RS.Open SQLAUX, DBSYSTEM, adOpenKeyset, adLockOptimistic
                If TRAS.ESCADENA Then
                    SQLAUX = "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE " & ModPlan.TRAS.FIELDK & "='" & LEmpresas.ListItems(X).Text & "'"
                Else
                    SQLAUX = "SELECT * FROM " & ModPlan.TRAS.TABLA & " WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                End If
                RSCOMPARA.Open SQLAUX, CNX_AUX, adOpenKeyset, adLockOptimistic
                If RS.RecordCount Then
                    FILATXT = RSCOMPARA.Fields(1)
                    If MsgBox("DESEA EL REEMPLAZAR EL REGISTRO " & FILATXT & " POR " & Me.LEmpresas.ListItems(X).ListSubItems(1).Text, vbYesNo, "CONFIRMAR") = vbYes Then
                        FLAG = 1
                        Select Case UCase(ModPlan.TRAS.TABLA)
                        Case "CCOSTOS"
                                SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "' ,RUC='" & RSCOMPARA.Fields(2) & "' ,FECHAING='" & RSCOMPARA.Fields(3) & "' ,CRONOGRAMA='" & RSCOMPARA.Fields(4) & "' WHERE CODCCOSTO='" & RSCOMPARA.Fields(0) & "'"
                        Case "AREASTRAB"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "' ,RUC='" & RSCOMPARA.Fields(2) & "' ,FECHAING='" & RSCOMPARA.Fields(3) & "' ,CRONOGRAMA='" & RSCOMPARA.Fields(4) & "' WHERE CODCCOSTO='" & RSCOMPARA.Fields(0) & "'"
                        Case "AFPS"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "' ,APOROBLI=" & RSCOMPARA.Fields(2) & " ,SEGURO=" & RSCOMPARA.Fields(3) & " ,TOPESEGURO=" & RSCOMPARA.Fields(4) & " ,COMISIONRA=" & RSCOMPARA.Fields(5) & " WHERE CODAFP='" & RSCOMPARA.Fields(0) & "'"
                        Case "CONCEPTOS"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "' ,TIPO=" & RSCOMPARA.Fields(2) & " ,FILA=" & RSCOMPARA.Fields(3) & " ,ESESCRITO=" & IIf(RSCOMPARA.Fields(4) = "VERDADERO", -1, 0) & " ,FORMULA='" & RSCOMPARA.Fields(5) & _
                            " ,COLPLANILLA='" & RSCOMPARA.Fields(6) & " ,MONEDA=" & RSCOMPARA.Fields(7) & " ,SUMAAFP=" & IIf(RSCOMPARA.Fields(8) = "VERDADERO", -1, 0) & ", SUMASALUD=" & IIf(RSCOMPARA.Fields(9) = "VERDADERO", -1, 0) & ", SUMAIES=" & IIf(RSCOMPARA.Fields(10) = "VERDADERO", -1, 0) & ", SUMARENTA=" & IIf(RSCOMPARA.Fields(11) = "VERDADERO", -1, 0) & ", SUMASCTR=" & IIf(RSCOMPARA.Fields(12) = "VERDADERO", -1, 0) & ", SUMACTS=" & IIf(RSCOMPARA.Fields(13) = "VERDADERO", -1, 0) & _
                            ", SUMAGRAT=" & IIf(RSCOMPARA.Fields(14) = "VERDADERO", -1, 0) & ", SUMAVAC=" & IIf(RSCOMPARA.Fields(15) = "VERDADERO", -1, 0) & ", SUMAT1=" & IIf(RSCOMPARA.Fields(16) = "VERDADERO", -1, 0) & _
                            ", SUMAT2=" & IIf(RSCOMPARA.Fields(17) = "VERDADERO", -1, 0) & ", SUMAT3=" & IIf(RSCOMPARA.Fields(18) = "VERDADERO", -1, 0) & ", SUMAT4=" & IIf(RSCOMPARA.Fields(19) = "VERDADERO", -1, 0) & ", SUMAT5=" & IIf(RSCOMPARA.Fields(20) = "VERDADERO", -1, 0) & ", TIPOINFO=" & RSCOMPARA.Fields(21) & ", TIPOREMU=" & RSCOMPARA.Fields(22) & ", TIPOQUINTA=" & RSCOMPARA.Fields(23) & ", ENLACE='" & RSCOMPARA.Fields(24) & ", ENLACE='" & RSCOMPARA.Fields(25) & _
                            ", FLAG=" & RSCOMPARA.Fields(26) & ", IMPRESIONFIJA=" & RSCOMPARA.Fields(27) & ", TIPOCTS=" & RSCOMPARA.Fields(28) & ", TIPOVAC=" & RSCOMPARA.Fields(29) & ", TIPOGRA=" & RSCOMPARA.Fields(30) & ", INDCTS=" & RSCOMPARA.Fields(31) & ", INDVAC=" & RSCOMPARA.Fields(32) & ", INDGRA=" & RSCOMPARA.Fields(33) & " WHERE CODAFP='" & RSCOMPARA.Fields(0) & "'"
                        Case "BANCOS"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "' WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "DOCUMENTOS"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET DESCRIP='" & RSCOMPARA.Fields(1) & "' WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "COLUMPL"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "', VALOR=" & RSCOMPARA.Fields(1) & "', INDICE=" & RSCOMPARA.Fields(2) & "', TIPO=" & RSCOMPARA.Fields(3) & "  WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "FORMULASCTS"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "', FORMULA='" & RSCOMPARA.Fields(2) & "', CRITERIO='" & RSCOMPARA.Fields(3) & "', TIPO=" & RSCOMPARA.Fields(4) & " WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "FORMULASGRATI"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "', FORMULA='" & RSCOMPARA.Fields(2) & "', CRITERIO='" & RSCOMPARA.Fields(3) & "', TIPO=" & RSCOMPARA.Fields(4) & " WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "FORMULASVAC"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "', FORMULA='" & RSCOMPARA.Fields(2) & "', CRITERIO='" & RSCOMPARA.Fields(3) & "', TIPO=" & RSCOMPARA.Fields(4) & " WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        Case "FORMULASUTIL"
                            SQLEXEC = "UPDATE " & ModPlan.TRAS.TABLA & " SET NOMBRE='" & RSCOMPARA.Fields(1) & "', FORMULA='" & RSCOMPARA.Fields(2) & "', CRITERIO='" & RSCOMPARA.Fields(3) & "' WHERE " & ModPlan.TRAS.FIELDK & "='" & RSCOMPARA.Fields(0) & "'"
                        End Select
                    End If
                Else
                    FLAG = 1
                    If ModPlan.TRAS.ESCADENA Then
                        SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT * FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                ModPlan.TRAS.TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "='" & LEmpresas.ListItems(X).Text & "'"
                    Else
                        If IDENTITY Then
                            If ExisteCampo("TIPO", ModPlan.TRAS.TABLA, DBSYSTEM) Then
                                TABLAIDEN = ModPlan.TRAS.TABLA & "(NOMBRE, FORMULA, CRITERIO, TIPO) "
                                SQLEXEC = "INSERT INTO " & TABLAIDEN & " SELECT NOMBRE, FORMULA, CRITERIO, TIPO FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                ModPlan.TRAS.TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                            Else
                                TABLAIDEN = ModPlan.TRAS.TABLA & "(NOMBRE, FORMULA, CRITERIO)"
                                SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT NOMBRE, FORMULA, CRITERIO FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                ModPlan.TRAS.TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                            End If
                        Else
                                SQLEXEC = "INSERT INTO " & ModPlan.TRAS.TABLA & " SELECT * FROM [" & ModPlan.TRAS.RUTABASE & "].dbo.[" & _
                                ModPlan.TRAS.TABLA & "] WHERE " & ModPlan.TRAS.FIELDK & "=" & LEmpresas.ListItems(X).Text
                        End If
                    End If
                If FLAG = 1 Then DBSYSTEM.Execute SQLEXEC
                End If
                RS.Close
                RSCOMPARA.Close
            End If
        Next X
        CNX_AUX.Close
    End If
    Screen.MousePointer = 1
    MsgBox "LA OPERACIÓN SE REALIZÓ SATISFACTORIAMENTE", vbInformation, "CONFIRMACIÓN"
End Sub

Private Sub Command1_Click()
    VGLFRM = 1
    Unload Me
    frEmpTr.Show 1
End Sub

Private Sub CHECK2_CLICK()
On Error Resume Next
If Check2.Value = 1 Then
    For X = 1 To Me.LEmpresas.ListItems.Count
        LEmpresas.ListItems(X).Checked = True
    Next
Else
    For X = 1 To Me.LEmpresas.ListItems.Count
        LEmpresas.ListItems(X).Checked = False
    Next
End If
End Sub


Private Sub FORM_LOAD()
    xEmpresa.Caption = ModPlan.TRAS.EMPRESA
End Sub




