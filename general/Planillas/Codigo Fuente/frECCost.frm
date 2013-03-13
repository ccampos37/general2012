VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frECCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edición de Centro de costo"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frECCost.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar"
      Height          =   345
      Left            =   3915
      TabIndex        =   9
      Top             =   2790
      Width           =   1065
   End
   Begin VB.CommandButton cmAdiciona 
      Caption         =   "A&dicionar"
      Height          =   345
      Left            =   3915
      TabIndex        =   10
      Top             =   2340
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGLista 
      Height          =   3510
      Left            =   105
      TabIndex        =   11
      Top             =   2355
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   6191
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Periodos de Pago Pendientes"
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
   Begin VB.CommandButton cmCancela 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3915
      TabIndex        =   13
      Top             =   5445
      Width           =   1065
   End
   Begin VB.CommandButton cmAcepta 
      Caption         =   "&Aceptar"
      Height          =   345
      Left            =   3915
      TabIndex        =   12
      Top             =   4980
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "Centro de Costo"
      Height          =   2115
      Left            =   105
      TabIndex        =   4
      Top             =   135
      Width           =   4875
      Begin VB.CheckBox Crono 
         Caption         =   "&Utilizar Cronograma de Pagos Predeterminado"
         Height          =   255
         Left            =   210
         TabIndex        =   14
         Top             =   1710
         Width           =   3675
      End
      Begin MSComCtl2.DTPicker xFecha 
         Height          =   285
         Left            =   1350
         TabIndex        =   3
         Top             =   1350
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36493
         MinDate         =   2
      End
      Begin AplisetControlText.Aplitext xRuc 
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   960
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         MaxLength       =   11
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Top             =   630
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   285
         Left            =   1350
         TabIndex        =   0
         Top             =   300
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         MaxLength       =   10
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1410
         Width           =   1020
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "C.C. Contable"
         Height          =   195
         Left            =   210
         TabIndex        =   7
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   390
         Width           =   495
      End
   End
End
Attribute VB_Name = "frECCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSNOMS As New ADODB.Recordset

Private Sub cmAcepta_Click()
    If VPTAREA = "NUEVO" Then
        'SI ES UN REGISTRO NUEVO
        If Not VERIFICACODIGO Then
            xCodigo.SetFocus
            Exit Sub
        End If
    End If
    If xNombre.Text = "" Then
        MsgBox "DEBE INGRESAR UN NOMBRE DESCRIPTIVO DEL CENTRO DE COSTO VÁLIDO", vbCritical
        xNombre.SetFocus
        Exit Sub
    End If
    If REGSISTEMA.scTieneStConta Then
        If Trim(xRuc.Text) = "" Then
            MsgBox "Debe ingresar el C.C. contable", vbExclamation
            xRuc.SetFocus
            Exit Sub
        End If
    End If
    If VPTAREA = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO CCOSTOS (CODCCOSTO,NOMBRE,RUC,FECHAING, CRONOGRAMA) SELECT '" & xCodigo.Text & "','" & xNombre.Text & "','" & "" & xRuc.Tag & "'," & DateSQL(xFecha.Value) & "," & Crono.Value
    Else
        DBSYSTEM.Execute "UPDATE CCOSTOS SET NOMBRE='" & xNombre.Text & "',RUC='" & xRuc.Text & "',FECHAING=" & DateSQL(xFecha.Value) & ",CRONOGRAMA=" & Crono.Value & " WHERE CODCCOSTO='" & xCodigo.Text & "'"
    End If
    Unload Me
End Sub

Private Sub CMADICIONA_Click()
    Dim RSNOMBOL As New ADODB.Recordset
    RSNOMBOL.Open "SELECT CODIGO, NOMBRE FROM NOMBOL WHERE CERRADO=0 ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    If RSNOMBOL.RecordCount = 0 Then
        MsgBox "EL SISTEMA NO HA PODIDO ENCONTRAR PERIODOS DE PAGOS DENTRO DE LA BASE DE DATOS DEL CRONOGRAMA DE PAGOS, O PUEDE SER QUE ESTOS SE HAYAN AGOTADO", vbCritical
        Set RSNOMBOL = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSNOMBOL
    frmComun.Show 1
    
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT * FROM FECHAPAGO WHERE CODNOMBOL=" & VGUTIL(1) & _
    " AND CODREF='" & xCodigo.Text & "' AND TIPOAC=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    If RSAUX.RecordCount > 0 Then
        MsgBox "EL REGISTRO YA SE ENCUENTRA INGRESADO"
        Set RSAUX = Nothing
        Exit Sub
    End If
    
    If VGUTIL(1) <> "" Then
        DBSYSTEM.Execute "INSERT INTO FECHAPAGO (CODREF,TIPOAC,CODNOMBOL) VALUES ('" & xCodigo.Text & "',1," & VGUTIL(1) & ")"
    End If
    Set RSNOMBOL = Nothing
    RSNOMS.Requery
    Set DGLista.DataSource = RSNOMS
End Sub

Private Sub CMCANCELA_Click()
    Unload Me
End Sub

Private Sub CMQUITAR_CLICK()
    If RSNOMS.RecordCount = 0 Or RSNOMS.EOF Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM FECHAPAGO WHERE ID_FECHAPAGO=" & RSNOMS!ID_FECHAPAGO
    RSNOMS.Requery
    Set DGLista.DataSource = RSNOMS
End Sub

Private Sub CRONO_Click()
    If VPTAREA = "NUEVO" Then Exit Sub
    If Crono.Value = 1 Then
        DGLista.Visible = True
        cmQuitar.Visible = True
        cmAdiciona.Visible = True
    Else
        DGLista.Visible = False
        cmQuitar.Visible = False
        cmAdiciona.Visible = False
    End If
End Sub

Private Sub FORM_ACTIVATE()
    If VPTAREA = "EDITAR" Then
        xCodigo.Locked = True
    Else 'SI ES NUEVO, ENTONCES DGLISTA NO DEBE MOSTRARSE
        DGLista.Visible = False
        cmQuitar.Visible = False
        cmAdiciona.Visible = False
        xFecha.Value = Date
    End If
End Sub

Private Sub FORM_LOAD()
    'CARGA DE DATOS DESDE EL PANEL ANTERIOR
    'UNA DE LAS CONDICIONES ES QUE EL ORIGEN DEBE TENER TODA LA INFORMACIÓN
    If VPTAREA = "EDITAR" Then
        If REGSISTEMA.scTieneStConta Then
            xRuc.Locked = True
         Else
            xRuc.Locked = False
        End If
        xCodigo.Text = frCostos.lvCCostos.SelectedItem.Text
        xNombre.Text = frCostos.lvCCostos.SelectedItem.SubItems(1)
        xRuc.Text = frCostos.lvCCostos.SelectedItem.SubItems(2)
        xFecha.Value = frCostos.lvCCostos.SelectedItem.SubItems(3)
        Dim RSAUX As New ADODB.Recordset
        RSAUX.Open "SELECT * FROM CCOSTOS WHERE CODCCOSTO='" & xCodigo.Text & "'", DBSYSTEM, adOpenStatic
        If RSAUX.EOF Then
            MsgBox "SE HA MODIFICADO O ELIMINADO EL REGISTRO, POR FAVOR VUELVA A INTENTAR", vbInformation
            Unload Me
        End If
        RSNOMS.Open "SELECT ID_FECHAPAGO, NOMBRE FROM FECHAPAGO, NOMBOL WHERE CODIGO=CODNOMBOL AND TIPOAC=1 AND CODREF='" & xCodigo.Text & "' ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
        Set DGLista.DataSource = RSNOMS
        With RSAUX
            If !CRONOGRAMA = 1 Then
                Crono.Value = 1
            Else
                Crono.Value = 0
            End If
            CRONO_Click
        End With
    End If
End Sub

Public Function VERIFICACODIGO() As Boolean
    'VERIFICA SI EL CÓDIGO ES VÁLIDO
    Dim X As Byte, xCad As String
    If frCostos.EXISTECODIGO(xCodigo.Text) Then
        MsgBox "EL CÓDIGO INGRESADO YA EXISTE"
        VERIFICACODIGO = False
        Exit Function
    End If
    'NO PUEDEN EMPEZAR POR PUNTO
    If Right(xCodigo.Text, 1) = "." Or Left(xCodigo.Text, 1) = "." Then
        MsgBox "EL CÓDIGO NO PUEDE EMPEZAR NI TERMINAR EN PUNTO, VERIFIQUE LOS DATOS DEL CÓDIGO DEL CENTRO DE COSTO", vbCritical
        VERIFICACODIGO = False
        Exit Function
    End If
    'RECORRIDO PARA SABER SI EXISTE EL PADRE DEL CÓDIGO DEL CENTRO DE COSTO
    If InStr(xCodigo.Text, ".") > 0 Then
        For X = 0 To Len(xCodigo.Text) - 1
            If Mid(xCodigo.Text, Len(xCodigo.Text) - X, 1) = "." Then
                xCad = Left(xCodigo.Text, Len(xCodigo.Text) - (X + 1))
                If Not frCostos.EXISTECODIGO(xCad) Then
                    MsgBox "LA REFERENCIA AL CENTRO DE COSTO SUPERIOR NO EXISTE. NO EXISTE NINGÚN CENTRO DE COSTO CON CÓDIGO " & xCad, vbCritical
                    VERIFICACODIGO = False
                    Exit Function
                End If
            End If
        Next
    End If
    VERIFICACODIGO = True
End Function

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSNOMS = Nothing
End Sub
Private Sub xRuc_Click()
    Dim RSCC As New ADODB.Recordset
    If REGSISTEMA.scTieneStConta Then
        RSCC.Open "SELECT CENCOST_CODIGO,CENCOST_DESCRIPCION FROM CENTRO_COSTOS", CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"), adOpenStatic, adLockReadOnly
        VGUTIL(1) = ""
        frmComun.CONECTAR RSCC
        frmComun.Show 1
        If VGUTIL(1) <> "" Then
            xRuc.Tag = VGUTIL(1)
            xRuc.Text = VGUTIL(1) & ":" & VGUTIL(2)
        End If
    End If
End Sub

Private Sub xRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        xRuc.Text = ""
    End If
End Sub
