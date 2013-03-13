VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form FrmAsisElect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Asistencia de Marcador Electronico"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmAsisElect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar Prog 
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   2505
      Visible         =   0   'False
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton CmdImportar 
      Caption         =   "&Importar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   75
      TabIndex        =   4
      Top             =   1875
      Width           =   1080
   End
   Begin VB.CommandButton CmdAbrir 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   4065
      MaskColor       =   &H00808080&
      Picture         =   "FrmAsisElect.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1380
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      Begin VB.Label Label2 
         Caption         =   "En este proceso se importara los datos de un archivo de texto a la base de datos del sistema."
         Height          =   495
         Left            =   645
         TabIndex        =   1
         Top             =   300
         Width           =   3750
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   45
         Picture         =   "FrmAsisElect.frx":0C0C
         Top             =   165
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   4005
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin AplisetControlText.Aplitext TxtRuta 
      Height          =   345
      Left            =   75
      TabIndex        =   3
      Top             =   1380
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   609
      Locked          =   -1  'True
      Text            =   ""
   End
   Begin VB.Label LbProg 
      Caption         =   "Importando los datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   2310
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label Label1 
      Caption         =   "Archivo a Importar"
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   1110
      Width           =   1920
   End
End
Attribute VB_Name = "FrmAsisElect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDABRIR_Click()
On Error GoTo ERRACT
    Dim FILENAMEASIS As String
    FILENAMEASIS = GetSetting(App.CompanyName, "PLANILLAS", "DIRFILEASIS", "")
    If FILENAMEASIS <> "" Then CDialog.FileName = FILENAMEASIS
    CDialog.Filter = "ARCHIVO DE TEXTO  |*.TXT"
    CDialog.ShowOpen
    TxtRuta.Text = CDialog.FileName
    If TxtRuta.Text = "" Then
        CmdImportar.Enabled = False
        Exit Sub
      Else:
      CmdImportar.Enabled = True
    End If
    Exit Sub
ERRACT:
    MsgBox ERR.Description
End Sub

Private Sub CMDIMPORTAR_Click()
    Dim LINEA As String
    Dim RSIMPORT As New ADODB.Recordset
    Dim RSAUX As New ADODB.Recordset
    Dim RSASIS As New ADODB.Recordset
    Dim FECHA As String
    Dim FLAG As Boolean
    Dim FLAGMSG As Boolean
    
On Error GoTo Err1
    
    If MsgBox("ESTA SEGURO QUE DESEA A PROCEDER A IMPORTAR LOS DATOS" + Chr(13) + _
              "UNA VEZ PROCESADO NO PODRA DESHACER LOS CAMBIOS", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    
    Open TxtRuta.Text For Input As #1
    If ExisteTablaSQL("TMPIMPORT", DBSTARPLAN) Then DBSTARPLAN.Execute "DROP TABLE TMPIMPORT"
    DBSTARPLAN.Execute "CREATE TABLE TMPIMPORT (FECHA DATETIME,CODTRAB VARCHAR(8),CODCONC VARCHAR(8), VALOR  Numeric(20,2) )"
    RSIMPORT.Open "TMPIMPORT", DBSTARPLAN, adOpenKeyset, adLockOptimistic
    Do
        Line Input #1, LINEA
        If Len(LINEA) > 32 Then
            MsgBox "EL FORMATO DE ARCHIVO NO CORRESPONDE CON LA ESTRUCTURA DE LA TABLA ASISTENCIA", vbExclamation
            Close #1
            Screen.MousePointer = 1
            Exit Sub
        End If
        RSIMPORT.AddNew
        FECHA = Mid(LINEA, 1, 8)
        RSIMPORT!FECHA = DateSerial(Val(Mid(FECHA, 5, 4)), Val(Mid(FECHA, 3, 2)), Val(Mid(FECHA, 1, 2)))
        RSIMPORT!CODTRAB = Mid(LINEA, 9, 6)
        RSIMPORT!CODCONC = Mid(LINEA, 15, 8)
        RSIMPORT!VALOR = Val(Mid(LINEA, 23, 10))
        RSIMPORT.Update
    Loop Until EOF(1)
    Close #1
    
    Dim MSG As VbMsgBoxResult
    RSASIS.Open "ASIS" & Year(Date), DBSYSTEM, adOpenKeyset, adLockOptimistic
    
    Set RSIMPORT = Nothing
    RSIMPORT.Open "TMPIMPORT", DBSTARPLAN, adOpenKeyset, adLockReadOnly
    If RSIMPORT.RecordCount = 0 Then Exit Sub
    
    LbProg.Visible = True: Prog.Visible = True
    Prog.Max = RSIMPORT.RecordCount + 1
    Prog.Min = 0: Prog.Value = 1
    Me.Height = 3105
    RSIMPORT.MoveFirst
    Do While Not RSIMPORT.EOF
        Set RSAUX = Nothing
        RSAUX.Open "SELECT * FROM TRABAJADORES WHERE CODTRAB='" & Trim(RSIMPORT!CODTRAB) & "'", DBSYSTEM, adOpenKeyset
        If RSAUX.RecordCount = 0 Then
            FrmDetImp.Rich.Text = FrmDetImp.Rich.Text + Chr(13) + Chr(10) & "*.- TRABAJADOR : """ & Trim(RSIMPORT!CODTRAB) & """----NO EXISTE"
            FLAG = True
            GoTo MOVE
        End If
        Set RSAUX = Nothing
        RSAUX.Open "SELECT * FROM CONCEPTOS WHERE CODIGO='" & Trim(RSIMPORT!CODCONC) & "'", DBSYSTEM, adOpenKeyset
        If RSAUX.RecordCount = 0 Then
            FrmDetImp.Rich.Text = FrmDetImp.Rich.Text + Chr(13) + Chr(10) & "*.- EL CONCEPTO : """ & Trim(RSIMPORT!CODCONC) & """ DEL TRABAJADOR : " & Trim(RSIMPORT!CODTRAB) & "----NO EXISTE"
            FLAG = True
            GoTo MOVE
        End If
        
        'VALIDANDO QUE NO REPITA LOS DATOS DE LA IMPORTACION
        Set RSAUX = Nothing
        RSAUX.Open "SELECT * FROM ASIS2000 WHERE CODTRAB='" & Trim(RSIMPORT!CODTRAB) & "' AND CONCEPTO='" & Trim(RSIMPORT!CODCONC) & "' AND DIA=" & DateSQL(RSIMPORT!FECHA), DBSYSTEM, adOpenKeyset
        If RSAUX.RecordCount > 0 Then
            If Not FLAGMSG Then
                MSG = MsgBox("REGISTRO ENCONTRADO :" & Trim(RSIMPORT!CODTRAB) & "," & Trim(RSIMPORT!CODCONC) & "," & Trim(RSIMPORT!FECHA) & _
                Chr(13) & "(SI) ACTUALIZA REGISTRO POR REGISTRO (NO) ACTUALIZA TODOS LOS REGISTROS", vbQuestion + vbYesNoCancel, "ACTUALIZAR REGISTRO(S)")
                Select Case MSG
                    Case vbYes: DBSYSTEM.Execute "DELETE FROM ASIS2000 WHERE CODTRAB='" & Trim(RSIMPORT!CODTRAB) & "' AND CONCEPTO='" & Trim(RSIMPORT!CODCONC) & "' AND DIA=" & DateSQL(RSIMPORT!FECHA)
                    Case vbNo
                        DBSYSTEM.Execute "DELETE FROM ASIS2000 WHERE CODTRAB='" & Trim(RSIMPORT!CODTRAB) & "' AND CONCEPTO='" & Trim(RSIMPORT!CODCONC) & "' AND DIA=" & DateSQL(RSIMPORT!FECHA)
                        FLAGMSG = True
                    Case vbCancel: GoTo MOVE
                End Select
              Else: DBSYSTEM.Execute "DELETE FROM ASIS2000 WHERE CODTRAB='" & Trim(RSIMPORT!CODTRAB) & "' AND CONCEPTO='" & Trim(RSIMPORT!CODCONC) & "' AND DIA=" & DateSQL(RSIMPORT!FECHA)
            End If
        End If
        RSASIS.AddNew
        RSASIS!CODTRAB = RSIMPORT!CODTRAB
        RSASIS!DIA = RSIMPORT!FECHA
        RSASIS!CONCEPTO = RSIMPORT!CODCONC
        RSASIS!VALOR = RSIMPORT!VALOR
        RSASIS.Update
        
MOVE:
        Prog.Value = Prog.Value + 1
        RSIMPORT.MoveNext
    Loop
    Screen.MousePointer = 1
    If Not FLAG Then
        MsgBox "PROCESO COMPLETADO SATISFACTORIAMENTE. " & Chr(13) & "EL REGISTRO DE ASISTENCIA SE ACTUALIZO EN LA BASE DE DATOS"
     Else
        MsgBox "PROCESO COMPLETADO CON ALGUNAS INCOSISTENCIAS. " & Chr(13) & "ALGUNOS REGISTROS NO SE ACTUALIZARON EN LA BASE DE DATOS"
        FrmDetImp.Show 1
    End If
    SaveSetting App.CompanyName, "PLANILLAS", "DIRFILEASIS", TxtRuta.Text
    
    Set RSIMPORT = Nothing
    Set RSAUX = Nothing
    Set RSASIS = Nothing
    LbProg.Visible = False: Prog.Visible = False
    Me.Height = 2715
    Exit Sub
Err1:
     MsgBox ERR.Description
     Exit Sub
End Sub


