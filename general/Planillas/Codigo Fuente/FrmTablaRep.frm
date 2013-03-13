VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form FrmTablRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Reportes - Planilla"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6795
   Icon            =   "FrmTablaRep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FraLista 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3135
      Left            =   60
      TabIndex        =   19
      Top             =   705
      Width           =   6645
      Begin MSComctlLib.ListView Lista 
         Height          =   3075
         Left            =   -15
         TabIndex        =   20
         Top             =   45
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   5424
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descripción Rep."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Boleta Remun."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Planilla de Remun."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cabecera de Planillas"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Activo"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Codigo"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "TipFmt"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Fijos"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   5595
      TabIndex        =   16
      Top             =   6285
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5850
      Top             =   135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTablaRep.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTablaRep.frx":065E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   2310
      Left            =   75
      ScaleHeight     =   2250
      ScaleWidth      =   6585
      TabIndex        =   0
      Top             =   3930
      Width           =   6645
      Begin VB.Frame Frame1 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   30
         TabIndex        =   21
         Top             =   1695
         Width           =   2640
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Fijos"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   45
            TabIndex        =   22
            Top             =   75
            Width           =   2310
         End
      End
      Begin VB.ComboBox xClaseBoleta 
         Height          =   315
         ItemData        =   "FrmTablaRep.frx":0982
         Left            =   2505
         List            =   "FrmTablaRep.frx":098F
         Style           =   2  'Dropdown List
         TabIndex        =   17
         ToolTipText     =   "Seleccione un tipo formato de impresión de boleta"
         Top             =   390
         Width           =   1875
      End
      Begin VB.CommandButton CmdMant 
         Caption         =   "&Modificar"
         Height          =   315
         Index           =   2
         Left            =   5355
         TabIndex        =   11
         Top             =   720
         Width           =   1125
      End
      Begin VB.CommandButton CmdMant 
         Caption         =   "&Eliminar"
         Height          =   315
         Index           =   1
         Left            =   5355
         TabIndex        =   10
         Top             =   405
         Width           =   1125
      End
      Begin VB.CommandButton CmdMant 
         Caption         =   "&Nuevo"
         Height          =   315
         Index           =   0
         Left            =   5355
         TabIndex        =   9
         Top             =   75
         Width           =   1125
      End
      Begin AplisetControlText.Aplitext xBolRem 
         Height          =   285
         Left            =   2190
         TabIndex        =   1
         Top             =   105
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   503
      End
      Begin AplisetControlText.Aplitext xPlanRem 
         Height          =   300
         Left            =   2190
         TabIndex        =   2
         Top             =   720
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
      End
      Begin AplisetControlText.Aplitext xCabPlan 
         Height          =   300
         Left            =   2190
         TabIndex        =   6
         Top             =   1035
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   529
      End
      Begin AplisetControlText.Aplitext xDescRep 
         Height          =   300
         Left            =   2190
         TabIndex        =   8
         Top             =   1365
         Width           =   4260
         _ExtentX        =   7514
         _ExtentY        =   529
      End
      Begin VB.Frame FramChk 
         BackColor       =   &H00808080&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   315
         Left            =   5160
         TabIndex        =   14
         Top             =   1590
         Width           =   1335
         Begin VB.CheckBox chkACt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Caption         =   "Activo"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   330
            TabIndex        =   15
            Top             =   105
            Width           =   945
         End
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Formato de Impresión Tipo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   435
         Width           =   1920
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del Reporte"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   90
         TabIndex        =   7
         Top             =   1425
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Boleta de Remuneraciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   120
         Width           =   2445
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Planilla de Remuneraciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   765
         Width           =   1980
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cabecera de Planillas"
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   90
         TabIndex        =   3
         Top             =   1080
         Width           =   1965
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   360
      Left            =   1065
      TabIndex        =   12
      Top             =   345
      Width           =   3075
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   1050
      TabIndex        =   13
      Top             =   345
      Width           =   3075
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   135
      Picture         =   "FrmTablaRep.frx":09B9
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   450
      Picture         =   "FrmTablaRep.frx":1283
      Top             =   15
      Width           =   480
   End
   Begin VB.Menu mnu 
      Caption         =   "Prueba"
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu Mnuactivar 
         Caption         =   "Activar"
      End
      Begin VB.Menu MnuNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu MnuEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu MnuModificar 
         Caption         =   "Modificar"
      End
   End
End
Attribute VB_Name = "FrmTablRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS2 As New ADODB.Recordset
Dim XITEM As ListItem
Dim SWMANT As Integer

Private Sub LIMPIARTEXT()
    xBolRem.Text = ""
    xCabPlan.Text = ""
    xDescRep.Text = ""
    xPlanRem.Text = ""
    xClaseBoleta.ListIndex = 0
    chkACt.Value = 0
    Check1.Value = 0
End Sub

Private Sub ACTIVARTEXT(FLAG As Boolean)
    xBolRem.Locked = Not FLAG
    xCabPlan.Locked = Not FLAG
    xDescRep.Locked = Not FLAG
    xPlanRem.Locked = Not FLAG
    xClaseBoleta.Locked = Not FLAG
    FramChk.Enabled = FLAG
    Frame1.Enabled = FLAG
End Sub

Private Sub MNUACTIVAR_Click()
    Dim CODIGO As String
    'VALIDANDO QUE EXISTA SOLO UN REPORTE ACTIVO POR USUARIO
    
    If Mnuactivar.Caption = "ACTIVAR" Then
        Dim RS2 As New ADODB.Recordset
        RS2.Open "SELECT * FROM TABLREP WHERE CODEMPRESA='" & REGSISTEMA.RUC & "' AND CODUSU='" & REGSISTEMA.USER & "' AND ACTIVO=-1", DBSTARPLAN, adOpenKeyset
        If RS2.RecordCount > 0 Then
            MsgBox "SOLO PUEDE EXISTIR UN GRUPO DE REPORTES ACTIVOS", vbExclamation
            Exit Sub
        End If
        DBSTARPLAN.Execute "UPDATE TABLREP SET ACTIVO=-1 WHERE CODIGO=" & Trim(Lista.SelectedItem.SubItems(5))
      Else:
        DBSTARPLAN.Execute "UPDATE TABLREP SET ACTIVO=0 WHERE CODIGO=" & Trim(Lista.SelectedItem.SubItems(5))
    End If
    CODIGO = Trim(Lista.SelectedItem.KEY)
    Call CARGARDATOS
    Dim INFOUND As ListItem
    Set INFOUND = Lista.FindItem(CODIGO, 2)
    INFOUND.EnsureVisible  ' DESPLAZA LISTVIEW PARA MOSTRAR EL LISTITEM HALLADO.
    INFOUND.Selected = True   ' SELECCIONA EL LISTITEM.
    ' DEVUELVE EL ENFOQUE AL CONTROL PARA VER LA SELECCIÓN.
    Lista.SetFocus
    Call REFRESCARDATOS
End Sub

Private Sub CMDMANT_Click(INDEX As Integer)
    Select Case INDEX
        Case 0:
            If CmdMant(0).Caption = "&Nuevo" Then
                ACTIVARTEXT True
                LIMPIARTEXT
                CmdMant(1).Enabled = False
                CmdMant(2).Caption = "&Cancelar"
                CmdMant(0).Caption = "&Grabar"
                SWMANT = 1
                xBolRem.SetFocus
                FraLista.Enabled = False
              Else:
                If VALIDAR Then Exit Sub
                If VALIDARREPORTES Then Exit Sub
                If SWMANT = 1 Then GRABAR (0)
                If SWMANT = 2 Then GRABAR (1)
                Call CANCELAR
                Call CARGARDATOS
                If Lista.ListItems.Count = 0 Then Exit Sub
                Call REFRESCARDATOS
            End If
        Case 1:
            If Lista.ListItems.Count = 0 Then Exit Sub
            If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR ESTE REGISTRO", vbQuestion + vbYesNo) = vbYes Then
                Call ELIMINAR
            End If
        Case 2:
            If CmdMant(2).Caption = "&Modificar" Then
                FraLista.Enabled = False
                CmdMant(1).Enabled = False
                CmdMant(2).Caption = "&Cancelar"
                CmdMant(0).Caption = "&Grabar"
                ACTIVARTEXT True
                SWMANT = 2
                xBolRem.SetFocus
              Else:
                Call LIMPIARTEXT
                Call CANCELAR
                If Lista.ListItems.Count = 0 Then Exit Sub
                Call REFRESCARDATOS
            End If
    End Select
End Sub
Private Sub CANCELAR()
    CmdMant(1).Enabled = True
    CmdMant(2).Caption = "&Modificar"
    CmdMant(0).Caption = "&Nuevo"
    ACTIVARTEXT False
    SWMANT = 0
    FraLista.Enabled = True
End Sub

Private Sub GRABAR(OP As Integer)
    Select Case OP
        Case 0:
            DBSTARPLAN.Execute "INSERT INTO TABLREP(CODEMPRESA,CODUSU,FILEBOLETA,FILEPLANILLA,FILEPLANCAB,DESCRIP,ACTIVO,TIPFMT,FIJO) " & _
                          "VALUES ('" & Trim(REGSISTEMA.RUC) & "','" & Trim(REGSISTEMA.USER) & "','" & Trim(xBolRem.Text) & "','" & Trim(xPlanRem.Text) & "','" & Trim(xCabPlan.Text) & "','" & Trim(xDescRep.Text) & "'," & IIf(chkACt.Value = 1, -1, 0) & "," & Str(xClaseBoleta.ListIndex) & " ," & Check1.Value & ")"
        Case 1:
            DBSTARPLAN.Execute "UPDATE TABLREP SET   " & _
                          "CODEMPRESA='" & Trim(REGSISTEMA.RUC) & "'," & _
                          "CODUSU='" & Trim(REGSISTEMA.USER) & "'," & _
                          "FILEBOLETA='" & Trim(xBolRem.Text) & "'," & _
                          "FILEPLANILLA='" & Trim(xPlanRem.Text) & "'," & _
                          "FILEPLANCAB='" & Trim(xCabPlan.Text) & "'," & _
                          "DESCRIP='" & Trim(xDescRep.Text) & "'," & _
                          "ACTIVO=" & IIf(chkACt.Value = 1, -1, 0) & "," & _
                          "TIPFMT=" & Str(xClaseBoleta.ListIndex) & _
                          ", FIJO=" & Check1.Value & _
                          " WHERE CODIGO=" & Lista.SelectedItem.SubItems(5)
       End Select
End Sub
Private Function VALIDAR() As Boolean
VALIDAR = True
    If Trim(xBolRem.Text) = "" Then
        MsgBox "TIENE QUE INGRESAR UN FORMATO DE BOLETA", vbExclamation
        xBolRem.SetFocus: Exit Function
    End If
    
    If Trim(xPlanRem.Text) = "" Then
        MsgBox "TIENE QUE INGRESAR UN FORMATO DE PLANILLA", vbExclamation
        xPlanRem.SetFocus: Exit Function
    End If
    
    If Trim(xCabPlan.Text) = "" Then
        MsgBox "TIENE QUE INGRESAR UN FORMATO DE PLANILLA CABECERA", vbExclamation
        xCabPlan.SetFocus: Exit Function
    End If
    
    If Trim(xDescRep.Text) = "" Then
        MsgBox "TIENE QUE INGRESAR LA DESCRIPCION", vbExclamation
        xDescRep.SetFocus: Exit Function
    End If
    
    'VALIDANDO QUE EXISTA SOLO UN REPORTE ACTIVO POR USUARIO
    If SWMANT = 1 Then
        If chkACt.Value = 1 Then
            Dim RS2 As New ADODB.Recordset
            RS2.Open "SELECT * FROM TABLREP WHERE CODEMPRESA='" & REGSISTEMA.RUC & "' AND CODUSU='" & REGSISTEMA.USER & "' AND ACTIVO=-1", DBSTARPLAN, adOpenKeyset
            If RS2.RecordCount > 0 Then
                MsgBox "SOLO PUEDE EXISTIR UN GRUPO DE REPORTES ACTIVOS", vbExclamation
                Exit Function
            End If
        End If
    End If
    VALIDAR = False
End Function

Private Sub ELIMINAR()
    DBSTARPLAN.Execute "DELETE FROM TABLREP WHERE CODIGO=" & Lista.SelectedItem.SubItems(5)
    Call LIMPIARTEXT
    Call CARGARDATOS
    Call REFRESCARDATOS
End Sub


Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub MNUELIMINAR_Click()
    If Lista.ListItems.Count = 0 Then Exit Sub
    If MsgBox("ESTA SEGURO QUE DESEA ELIMINAR ESTE REGISTRO", vbQuestion + vbYesNo) = vbYes Then
        Call ELIMINAR
    End If
End Sub

Private Sub FORM_LOAD()
    Call LIMPIARTEXT
    Call CARGARDATOS
    SWMANT = 0
    ACTIVARTEXT False
    If Lista.ListItems.Count = 0 Then Exit Sub
    Call REFRESCARDATOS
End Sub
Private Sub CARGARDATOS()
    If RS2.State <> 0 Then RS2.Close
    RS2.Open "SELECT * FROM TABLREP WHERE CODEMPRESA='" & REGSISTEMA.RUC & "' AND CODUSU='" & REGSISTEMA.USER & "'", DBSTARPLAN
    Lista.ListItems.Clear
    Do While Not RS2.EOF
        Set XITEM = Lista.ListItems.Add(, "C" & Trim(Str(RS2!CODIGO)), RS2!DESCRIP, , 2)
        XITEM.SubItems(1) = RS2!FILEBOLETA
        XITEM.SubItems(2) = RS2!FILEPLANILLA
        XITEM.SubItems(3) = RS2!FILEPLANCAB
        XITEM.SubItems(4) = IIf(RS2!ACTIVO = -1, "X", "")
        XITEM.SubItems(5) = RS2!CODIGO
        XITEM.SubItems(6) = RS2!TipFmt
        XITEM.SubItems(7) = RS2!FIJO
        XITEM.Tag = XITEM.KEY
        RS2.MoveNext
    Loop
    Lista.Refresh
End Sub
Private Sub REFRESCARDATOS()
    If Lista.ListItems.Count = 0 Then Exit Sub
    With Lista.SelectedItem
        xDescRep.Text = .Text
        xBolRem.Text = .SubItems(1)
        xPlanRem.Text = .SubItems(2)
        xCabPlan.Text = .SubItems(3)
        chkACt.Value = IIf(.SubItems(4) = "X", 1, 0)
        Check1.Value = IIf(.SubItems(7) = True, 1, 0)
        xClaseBoleta.ListIndex = Val(.SubItems(6))
    End With
End Sub
 
Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RS2 = Nothing
    frSetup.MOSTRARREPACTIVO
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call REFRESCARDATOS
End Sub

Private Sub LISTA_MOUSEDOWN(Button As Integer, SHIFT As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Lista.SelectedItem.SubItems(4) = "X" Then
           Mnuactivar.Caption = "DESACTIVAR"
          Else: Mnuactivar.Caption = "ACTIVAR"
        End If
        PopupMenu mnu
    End If
End Sub

Private Sub MNUMODIFICAR_Click()
    FraLista.Enabled = False
    CmdMant(1).Enabled = False
    CmdMant(2).Caption = "&CANCELAR"
    CmdMant(0).Caption = "&GRABAR"
    ACTIVARTEXT True
    SWMANT = 2
    xBolRem.SetFocus
End Sub

Private Sub MNUNUEVO_Click()
    ACTIVARTEXT True
    LIMPIARTEXT
    CmdMant(1).Enabled = False
    CmdMant(2).Caption = "&CANCELAR"
    CmdMant(0).Caption = "&GRABAR"
    SWMANT = 1
    xBolRem.SetFocus
    FraLista.Enabled = False
End Sub

Private Sub XBOLREM_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case ",": KeyAscii = 0
        Case "'": KeyAscii = 0
    End Select
    KeyAscii = Asc((UCase(Chr(KeyAscii))))
End Sub

Private Function VALIDARREPORTES() As Boolean
    VALIDARREPORTES = True
    
    'VALIDANDO SI ESTA EL ARCHIVO DE FORMATO DE BOLETAS
    If UCase(Dir$(REGSISTEMA.REPORTES & Trim(xBolRem.Text))) <> UCase(Trim(xBolRem.Text)) Then
        MsgBox "ESE FORMATO DE BOLETA :" & xBolRem.Text & " NO SE ENCUENTRA EN EL REGISTRO DE REPORTES", vbExclamation
        xBolRem.SetFocus
        Exit Function
    End If
    
    'VALIDANDO SI ESTA EL ARCHIVO DE FORMATO DE PLANILLAS
    If UCase(Dir$(REGSISTEMA.REPORTES & Trim(xPlanRem.Text))) <> UCase(Trim(xPlanRem.Text)) Then
        MsgBox "ESE FORMATO DE PLANILLA :" & xPlanRem.Text & ", NO SE ENCUENTRA EN EL REGISTRO DE REPORTES", vbExclamation
        xPlanRem.SetFocus
        Exit Function
    End If
    
    'VALIDANDO SI ESTA EL ARCHIVO DE FORMATO DE PLANILLAS CABECERAS
    If UCase(Dir$(REGSISTEMA.REPORTES & Trim(xCabPlan.Text))) <> UCase(Trim(xCabPlan.Text)) Then
        MsgBox "ESE FORMATO DE PLANILLA CABECERA :" & xCabPlan.Text & ", NO SE ENCUENTRA EN EL REGISTRO DE REPORTES", vbExclamation
        xCabPlan.SetFocus
        Exit Function
    End If
    VALIDARREPORTES = False
End Function

Private Sub XCABPLAN_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case ",": KeyAscii = 0
        Case "'": KeyAscii = 0
    End Select
    KeyAscii = Asc((UCase(Chr(KeyAscii))))
End Sub

Private Sub XDESCREP_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case ",": KeyAscii = 0
        Case "'": KeyAscii = 0
    End Select
    KeyAscii = Asc((UCase(Chr(KeyAscii))))
End Sub

Private Sub XPLANREM_KeyPress(KeyAscii As Integer)
    Select Case Chr(KeyAscii)
        Case ",": KeyAscii = 0
        Case "'": KeyAscii = 0
    End Select
    KeyAscii = Asc((UCase(Chr(KeyAscii))))
End Sub


