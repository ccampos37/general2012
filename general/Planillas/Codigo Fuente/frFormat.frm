VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frFormatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formatos de Planilla"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6615
   Icon            =   "frFormat.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   6615
   Tag             =   "Panel de Diseño de Formatos de Boletas"
   Begin VB.Frame Frame1 
      Caption         =   "Identificación del Formato"
      Height          =   1275
      Left            =   135
      TabIndex        =   2
      Top             =   60
      Width           =   5100
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   945
         TabIndex        =   10
         Top             =   345
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.CommandButton xAbrir 
         Height          =   465
         Left            =   3870
         Picture         =   "frFormat.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Abrir formato de planilla"
         Top             =   705
         Width           =   465
      End
      Begin VB.CommandButton cmNuevos 
         Height          =   465
         Left            =   4395
         Picture         =   "frFormat.frx":0C0C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Nuevo Formato de Planilla"
         Top             =   705
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Quitar Concepto"
         Height          =   195
         Left            =   2655
         TabIndex        =   6
         ToolTipText     =   "Eliminar el Concepto de Remuneraciones seleccionado del Formato de Planilla"
         Top             =   840
         Width           =   1155
      End
      Begin VB.Image QuitaCnpt 
         Height          =   240
         Left            =   2355
         Picture         =   "frFormat.frx":104E
         ToolTipText     =   "Eliminar el Concepto de Remuneraciones seleccionado del Formato de Planilla"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Agregar Concepto"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         ToolTipText     =   "Agregar un Concepto de Remuneraciones al Formato de Planilla"
         Top             =   840
         Width           =   1290
      End
      Begin VB.Image AgregaCnpt 
         Height          =   240
         Left            =   285
         Picture         =   "frFormat.frx":1390
         ToolTipText     =   "Agregar un Concepto de Remuneraciones al Formato de Planilla"
         Top             =   810
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   255
         TabIndex        =   3
         Top             =   390
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView xLista 
      Height          =   4095
      Left            =   135
      TabIndex        =   1
      Top             =   1380
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Planilla"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Formula"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   3675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":16D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":1A26
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":2302
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":2BDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":3A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frFormat.frx":3D56
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frFormat.frx":40AA
      ToolTipText     =   "Eliminar todo el Formato de Planilla"
      Top             =   5490
      Width           =   480
   End
   Begin VB.Image xGuardar 
      Height          =   240
      Left            =   1935
      Picture         =   "frFormat.frx":44EC
      ToolTipText     =   "Guerdar el Formato de Planilla"
      Top             =   5580
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Formatos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   5430
      TabIndex        =   0
      ToolTipText     =   "Programado por Daniel Yafac Baquedano: danielyafac@hotmail.com para Enterprise Solutions S.A."
      Top             =   660
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "frFormat.frx":482E
      ToolTipText     =   "Programado por Daniel Yafac Baquedano: danielyafac@hotmail.com para Enterprise Solutions S.A."
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "Guardar Formato "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2310
      TabIndex        =   4
      ToolTipText     =   "Guerdar el Formato de Planilla"
      Top             =   5603
      Width           =   1230
   End
   Begin VB.Label LblEliminar 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Eliminar Formato "
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      ToolTipText     =   "Eliminar todo el Formato de Planilla"
      Top             =   5633
      Width           =   1200
   End
End
Attribute VB_Name = "frFormatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REGACT As REGWIN

Private Sub AGREGACNPT_Click()
    If xNombre.Text = "" Or xNombre.Tag = "" Then
        MsgBox "Falta seleccionar un formato de planilla", vbCritical
        Exit Sub
    End If
    Dim RSCONCEP As New ADODB.Recordset
    RSCONCEP.Open "SELECT CODIGO, NOMBRE,COMENTARIO,TIPO,FORMULA,COLPLANILLA FROM CONCEPTOS WHERE CODIGO<>'REDONDEO' AND CODIGO NOT IN (SELECT CONCEPTO FROM FORMARUBS WHERE ID_FORMATO=" & xNombre.Tag & ") ORDER BY NOMBRE", DBSYSTEM, adOpenKeyset, adLockOptimistic
    frmComun.CONECTAR RSCONCEP
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        AgregaCnpt.Tag = VGUTIL(1)
    Else
        Set RSCONCEP = Nothing
        Exit Sub
    End If
    If Len(VGUTIL(1) >= 7) Then
        If Mid(VGUTIL(1), 1, 2) = "XX" Then
            MsgBox "Concepto reservado del sistema no se puede agregar al formato", vbExclamation
            Exit Sub
        End If
    End If
    On Error GoTo ERRYAEXISTE
    Dim XLIST As ListItem
    Set XLIST = xLista.ListItems(AgregaCnpt.Tag)
    Set RSCONCEP = Nothing
    Exit Sub
ERRYAEXISTE:
    'SE INGRESA A ESTA PARTE SI EL CODIGO NO ESTÁ EN LA LISTA
    With RSCONCEP
        Set XLIST = xLista.ListItems.Add(, !Codigo, !Codigo, , !TIPO + 1)
        XLIST.SubItems(1) = !NOMBRE
        XLIST.SubItems(2) = !COLPLANILLA
        XLIST.SubItems(3) = !FORMULA
        .MoveNext
    End With
    Resume Next
End Sub

Private Sub CMNUEVOS_Click()
    Dim STRCAD As String
    STRCAD = InputBox("ESCRIBA EL NOMBRE PARA EL NUEVO FORMATO DE PLANILLA:", "MARFICE PLANILLAS")
    If STRCAD = "" Then
        Exit Sub
    End If
    If Len(STRCAD) >= 50 Then
        MsgBox "El nombre del Formato de Planilla no debe exceder de 50 caracteres. La información no fue grabada", vbInformation
        Exit Sub
    End If
    DBSYSTEM.Execute "INSERT INTO FORMATOS (NOMBRE,FECHAING) VALUES ('" & STRCAD & "'," & DateSQL(Date) & ")"
End Sub

Private Sub LABEL5_Click()
XABRIR_Click
End Sub

Private Sub IMAGE2_Click()
LBLELIMINAR_Click
End Sub

Private Sub LABEL4_Click()
    If xNombre.Text = "" Or xNombre.Tag = "" Then
        MsgBox "Falta seleccionar un Formato de Planilla", vbCritical
        Exit Sub
    End If
    XGUARDAR_Click
End Sub

Private Sub LABEL7_Click()
AGREGACNPT_Click
End Sub

Private Sub LABEL8_Click()
QUITACNPT_Click
End Sub

Private Sub LBLELIMINAR_Click()
    Dim XNUM As Integer
    If MsgBox("Realmente desea eliminar el Formato de Planilla: " & xNombre.Text, vbYesNo) = vbNo Then Exit Sub
    DBSYSTEM.Execute "DELETE FROM FORMATOS WHERE ID_FORMATO=" & xNombre.Tag, XNUM
    If XNUM = 0 Then
        MsgBox "No se realizaron los cambios, puede deberse a que otro usuario ha eliminado o tratado de leiminar el Formato de Planilla", vbCritical
        Exit Sub
    Else
        MsgBox "El elemento se ha eliminado de la Base de Datos. Imposible volver a recuperar el registro", vbInformation, "TAREA COMPLETADA"
        xNombre.Text = ""
        xNombre.Tag = ""
        xLista.ListItems.Clear
    End If
End Sub

Private Sub QUITACNPT_Click()
    On Error GoTo ERRQUITAR
    xLista.ListItems.Remove xLista.SelectedItem.KEY
    Exit Sub
ERRQUITAR:
    Beep
    Resume Next
End Sub

Private Sub XABRIR_Click()
    Dim STRCAD As String
    Dim RSFORMA As New ADODB.Recordset
    RSFORMA.Open "SELECT * FROM FORMATOS ORDER BY NOMBRE", DBSYSTEM, adOpenStatic
    If RSFORMA.EOF Or RSFORMA.RecordCount = 0 Then
        MsgBox "No se ha encontrado Formatos de Planilla almacenados en la Base de Datos del Sistema. Por favor cree uno nuevo", vbCritical
        Set RSFORMA = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSFORMA
    frmComun.Show 1
    If VGUTIL(1) = "" Then
        MsgBox "Acción cancelada por el Usuario del Sistema", vbInformation
        Set RSFORMA = Nothing
        Exit Sub
    End If
    xLista.ListItems.Clear
    xNombre.Text = RSFORMA!NOMBRE
    xNombre.Tag = RSFORMA!ID_FORMATO
    Set RSFORMA = Nothing
    Dim XNODE As ListItem
    STRCAD = "SELECT CODIGO,NOMBRE,CONCEPTOS.TIPO,FORMULA,COLPLANILLA FROM CONCEPTOS, FORMARUBS WHERE FORMARUBS.CONCEPTO=CONCEPTOS.CODIGO AND FORMARUBS.ID_FORMATO=" & xNombre.Tag & " ORDER BY CONCEPTOS.TIPO, NOMBRE"
    RSFORMA.Open STRCAD, DBSYSTEM, adOpenStatic, adLockOptimistic
    With RSFORMA
    Do While Not .EOF
        Set XNODE = xLista.ListItems.Add(, !Codigo, !Codigo, , !TIPO + 1)
        XNODE.SubItems(1) = !NOMBRE
        XNODE.SubItems(2) = !COLPLANILLA
        XNODE.SubItems(3) = !FORMULA
        .MoveNext
    Loop
    End With
    Set RSFORMA = Nothing
    Set XNODE = Nothing
    xLista.ColumnHeaders(2).Width = 3360.189
End Sub

Private Sub XGUARDAR_Click()
    If xNombre.Text = "" Or xNombre.Tag = "" Then
        MsgBox "Falta seleccionar un Formato de Planilla", vbCritical
        Exit Sub
    End If
    If xLista.ListItems.Count = 0 Then
        MsgBox "No existe nada por garbar. Si ha existido anteriormente un Formato para el Centro de Costo y del tipo especificado, este se ha eliminado", vbCritical
        Exit Sub
    End If
    Dim STRCAD As String, XITEM As ListItem
    STRCAD = "DELETE FROM FORMARUBS WHERE ID_FORMATO=" & xNombre.Tag
    DBSYSTEM.Execute STRCAD
    For Each XITEM In xLista.ListItems
        DBSYSTEM.Execute "INSERT INTO FORMARUBS (ID_FORMATO,CONCEPTO) VALUES (" & xNombre.Tag & ",'" & XITEM.Text & "')"
    Next
    Set XITEM = Nothing
    MsgBox "Se guardarón los datos satisfactoriamente", vbInformation
End Sub

Public Sub COMANDOTOOLBAR(COMANDO As String)
    Select Case UCase(COMANDO)
        Case "NUEVO"
            AGREGACNPT_Click
        Case "EDITAR"
            QUITACNPT_Click
        Case Else
            MsgBox "Modulo no disponible"
    End Select
End Sub

Private Sub XLISTA_DblClick()
    If xLista.ListItems.Count = 0 Then Exit Sub
    VPTAREA = "EDITAR"
    VPCODTMP = xLista.SelectedItem.KEY
    frECnpt.Show 1
End Sub

Private Sub XNOMBRE_DblClick()
    XABRIR_Click
End Sub

Private Sub XRECARGAR_Click()
XABRIR_Click
End Sub

