VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frFormulasCTS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmulas de C.T.S."
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frFormulasCTS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   5775
      TabIndex        =   5
      Top             =   3330
      Width           =   1020
   End
   Begin VB.CommandButton cmBorrar 
      Caption         =   "&Borrar"
      Height          =   315
      Left            =   2475
      TabIndex        =   4
      Top             =   3330
      Width           =   1020
   End
   Begin VB.CommandButton cmEditar 
      Caption         =   "&Editar"
      Height          =   315
      Left            =   1320
      TabIndex        =   3
      Top             =   3330
      Width           =   1020
   End
   Begin VB.CommandButton cmNuevo 
      Caption         =   "&Nuevo"
      Height          =   315
      Left            =   165
      TabIndex        =   2
      Top             =   3330
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Caption         =   "Descripción de Fórmula de CTS"
      Enabled         =   0   'False
      Height          =   2430
      Left            =   75
      TabIndex        =   1
      Top             =   3795
      Width           =   6810
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   225
         Left            =   255
         TabIndex        =   21
         Top             =   1770
         Width           =   2505
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "General"
         Height          =   285
         Left            =   240
         TabIndex        =   20
         Top             =   2025
         Width           =   2520
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2145
         TabIndex        =   19
         Top             =   720
         Width           =   330
      End
      Begin MSComctlLib.ListView Lv 
         Height          =   825
         Left            =   255
         TabIndex        =   16
         Top             =   2400
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   1455
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Descrip"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "descri"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox Text1 
         Height          =   330
         Left            =   195
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   2340
         Visible         =   0   'False
         Width           =   3885
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   5655
         Top             =   1200
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
               Picture         =   "frFormulasCTS.frx":08CA
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   5625
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   4365
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ComboBox xTipo 
         Height          =   315
         ItemData        =   "frFormulasCTS.frx":0BE6
         Left            =   2550
         List            =   "frFormulasCTS.frx":0BF0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1350
         Width           =   2715
      End
      Begin AplisetControlText.Aplitext xCriterio 
         Height          =   285
         Left            =   2535
         TabIndex        =   11
         Top             =   1020
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   503
         MaxLength       =   240
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xFormula 
         Height          =   285
         Left            =   2535
         TabIndex        =   9
         Top             =   690
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   503
         MaxLength       =   240
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   285
         Left            =   2535
         TabIndex        =   7
         Top             =   360
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   503
         MaxLength       =   50
         Text            =   ""
         SinBlancos      =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   180
         Left            =   225
         TabIndex        =   18
         Top             =   2310
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Remuneración"
         Height          =   195
         Left            =   270
         TabIndex        =   12
         Top             =   1410
         Width           =   1620
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Criterio de Cómputo"
         Height          =   195
         Left            =   285
         TabIndex        =   10
         Top             =   1065
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fórmula de Acción"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   735
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre de Rem. Computable"
         Height          =   195
         Left            =   270
         TabIndex        =   6
         Top             =   405
         Width           =   2085
      End
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3105
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   5477
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "Fórmulas de C.T.S."
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "Nombre"
         Caption         =   "Remuneraciones Computables"
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
         DataField       =   "Formula"
         Caption         =   "Formula de Acción"
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
      BeginProperty Column02 
         DataField       =   "Tipo"
         Caption         =   "Tipo"
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
      BeginProperty Column03 
         DataField       =   "Criterio"
         Caption         =   "Criterio"
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
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1904.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   599.811
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   3705
      Left            =   75
      Top             =   60
      Width           =   6795
   End
End
Attribute VB_Name = "frFormulasCTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSFORMS As ADODB.Recordset
Dim POS As Integer

Private Sub CMACEPTAR_CLICK()
    'ON ERROR RESUME NEXT
    If XNombre.Text = "" Then
        MsgBox "Falta especificar un nombre", vbInformation
        XNombre.SetFocus
        Exit Sub
    End If
    If XFormula.Text = "" Then
        MsgBox "Falta especificar una formula", vbInformation
        XFormula.SetFocus
        Exit Sub
    End If
    If Frame1.Tag = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO FORMULASCTS (NOMBRE,FORMULA,CRITERIO,TIPO, AFECTOPRO, GENE ) VALUES ('" & XNombre.Text & "','" & XFormula.Text & "','" & xCriterio.Text & "'," & xTipo.ListIndex & ", " & Check1.Value & ", " & Check2.Value & ")"
    Else
        DBSYSTEM.Execute "UPDATE FORMULASCTS SET NOMBRE='" & XNombre.Text & "', FORMULA='" & XFormula.Text & "', CRITERIO='" & xCriterio.Text & "', TIPO=" & xTipo.ListIndex & ", AFECTOPRO=" & Check1.Value & ", GENE=" & Check2.Value & " WHERE CODIGO=" & RSFORMS!Codigo
    End If
    RSFORMS.Requery
    Set xData.DataSource = RSFORMS
    Frame1.Enabled = False
    cmAceptar.Visible = False
    cmCancelar.Visible = False
    XDATA_ROWCOLCHANGE 0, 0
    OCULTAR
End Sub

Private Sub CMBORRAR_CLICK()
    On Error Resume Next
    If RSFORMS.EOF Or RSFORMS.BOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If MsgBox("Esta seguro de eliminar la formula de C.T.S. :" & RSFORMS!NOMBRE, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    RSFORMS.Delete
    Call Limpiar
    If RSFORMS.RecordCount > 0 Then RSFORMS.MoveFirst
    XDATA_ROWCOLCHANGE 0, 0
End Sub

Private Sub CMCANCELAR_CLICK()
    Frame1.Enabled = False
    cmAceptar.Visible = False
    cmCancelar.Visible = False
    OCULTAR
    XDATA_ROWCOLCHANGE 0, 0
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMEDITAR_CLICK()
    If RSFORMS.EOF Then Exit Sub
    Frame1.Tag = "EDITAR"
    Frame1.Enabled = True
    OCULTAR
    XNombre.SetFocus
End Sub

Private Sub CMNUEVO_CLICK()
    Frame1.Enabled = True
    Frame1.Tag = "NUEVO"
    OCULTAR
    Call Limpiar
    XNombre.SetFocus
End Sub
Private Sub Limpiar()
    XNombre.Text = ""
    XFormula.Text = ""
    xCriterio.Text = ""
    xTipo.ListIndex = 0
    Check1.Value = 0
    Check2.Value = 0
End Sub

Private Sub Form_Load()
    Set RSFORMS = New ADODB.Recordset
    RSFORMS.Open "FORMULASCTS", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSFORMS
    xTipo.ListIndex = 0
    MOVECOMPLETE
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSFORMS = Nothing
End Sub

Public Sub OCULTAR()
    If Frame1.Enabled Then
        cmAceptar.Visible = True
        cmCancelar.Visible = True
        cmBorrar.Enabled = False
        cmCerrar.Enabled = False
        cmEditar.Enabled = False
        cmNuevo.Enabled = False
        xData.Enabled = False
    Else
        cmAceptar.Visible = False
        cmCancelar.Visible = False
        cmBorrar.Enabled = True
        cmCerrar.Enabled = True
        cmEditar.Enabled = True
        cmNuevo.Enabled = True
        xData.Enabled = True
    End If
End Sub

Private Sub MOVECOMPLETE()
    On Error Resume Next
    If RSFORMS.EOF Or RSFORMS.BOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If Frame1.Enabled Then Exit Sub
    XNombre.Text = RSFORMS!NOMBRE
    XFormula.Text = RSFORMS!FORMULA
    xTipo.ListIndex = RSFORMS!TIPO
    xCriterio.Text = RSFORMS!CRITERIO
    Check1.Value = IIf(RSFORMS!AFECTOPRO, 1, 0)
    Check2.Value = IIf(RSFORMS!GENE, 1, 0)
End Sub

Private Sub LV_KEYPRESS(KeyAscii As Integer)
    Dim I As Integer
    Dim CAD1 As String, CAD2 As String
    If KeyAscii = 27 Then Lv.Visible = False
    If KeyAscii = 13 Then
        CAD1 = Mid(Text1.Text, 1, POS)
        CAD2 = Mid(Text1.Text, POS + 1, Len(Text1.Text) - POS)
        Text1.Text = CAD1 + Lv.SelectedItem.Text + CAD2
        Text1.SelStart = POS + Len(Lv.SelectedItem.Text)
        Lv.Visible = False
    End If
End Sub

Private Sub LV_LOSTFOCUS()
    Lv.Visible = False
End Sub

Private Sub TEXT1_KEYDOWN(KEYCODE As Integer, Shift As Integer)
    If KEYCODE = 70 And Shift = 2 Then
        Dim CnAux As New ADODB.Connection
        With CnAux
         .CursorLocation = adUseClient
         .Provider = "MICROSOFT.JET.OLEDB.3.51"
         .ConnectionString = "DATA SOURCE=" & REGSISTEMA.ARCHIVOWENTPL
         .Open
        End With
        Dim RSTABFUN As New ADODB.Recordset
        Dim XLIST As ListItem
        RSTABFUN.Open "TABFUN", CnAux
        Lv.ListItems.Clear
        RSTABFUN.MoveFirst
        Do While Not RSTABFUN.EOF
          Set XLIST = Lv.ListItems.Add(, "C" & Trim(Str(RSTABFUN!Codigo)), RSTABFUN!FUNCION, , 1)
          XLIST.SubItems(1) = RSTABFUN!DESCRI
          RSTABFUN.MoveNext
        Loop
        Lv.Left = Lv.Left + (POS * 15)
        Lv.Visible = True
        Lv.SetFocus
   End If
End Sub

Private Sub TEXT1_KEYUP(KEYCODE As Integer, Shift As Integer)
    POS = Text1.SelStart
    Label5.Caption = Str(POS)
End Sub

Private Sub xCriterio_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub

Private Sub XDATA_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
MOVECOMPLETE
End Sub

Private Sub XFormula_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub
Private Sub xNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub

