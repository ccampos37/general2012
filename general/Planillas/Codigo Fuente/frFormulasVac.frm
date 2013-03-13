VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frFormulasVac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fórmulas de Vacaciones"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frFormulasVac.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
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
      Caption         =   "Descripción de Fórmula de Vacaciones"
      Enabled         =   0   'False
      Height          =   2475
      Left            =   60
      TabIndex        =   1
      Top             =   3780
      Width           =   6810
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "Activo"
         Height          =   225
         Left            =   240
         TabIndex        =   18
         Top             =   1755
         Width           =   2490
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "General"
         Height          =   285
         Left            =   225
         TabIndex        =   17
         Top             =   2055
         Width           =   2505
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2175
         TabIndex        =   16
         Top             =   705
         Width           =   285
      End
      Begin AplisetControlText.Aplitext XFormula 
         Height          =   315
         Left            =   2490
         TabIndex        =   14
         Top             =   660
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xNombre 
         Height          =   300
         Left            =   2490
         TabIndex        =   13
         Top             =   345
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   529
         Text            =   ""
      End
      Begin VB.CommandButton cmCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   5640
         TabIndex        =   7
         Top             =   2070
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmAceptar 
         Caption         =   "&Aceptar"
         Height          =   315
         Left            =   4380
         TabIndex        =   9
         Top             =   2070
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ComboBox xTipo 
         Height          =   315
         ItemData        =   "frFormulasVac.frx":08CA
         Left            =   2490
         List            =   "frFormulasVac.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1350
         Width           =   2715
      End
      Begin AplisetControlText.Aplitext xCriterio 
         Height          =   315
         Left            =   2490
         TabIndex        =   15
         Top             =   990
         Width           =   4080
         _ExtentX        =   7197
         _ExtentY        =   556
         Text            =   ""
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
      Caption         =   "Fórmulas de Vacaciones"
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
Attribute VB_Name = "frFormulasVac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSFORMS As ADODB.Recordset

Private Sub CMACEPTAR_CLICK()
    If xNombre.Text = "" Then
        MsgBox "Falta especificar un nombre", vbInformation
        xNombre.SetFocus
        Exit Sub
    End If
    If xFormula.Text = "" Then
        MsgBox "Falta especificar una formula", vbInformation
        xFormula.SetFocus
        Exit Sub
    End If
    If Frame1.Tag = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO FORMULASVAC (NOMBRE,FORMULA,CRITERIO,TIPO,AFECTOPRO,GENE) VALUES ('" & xNombre.Text & "','" & xFormula.Text & "','" & xCriterio.Text & "'," & xTipo.ListIndex & "," & Check1.Value & "," & Check2.Value & ")"
    Else
        DBSYSTEM.Execute "UPDATE FORMULASVAC SET NOMBRE='" & xNombre.Text & "', FORMULA='" & xFormula.Text & "',CRITERIO='" & xCriterio.Text & "',TIPO=" & xTipo.ListIndex & ", AFECTOPRO=" & Check1.Value & ", GENE=" & Check2.Value & " WHERE CODIGO=" & RSFORMS!Codigo
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
    If RSFORMS.EOF Or RSFORMS.BOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If MsgBox("Esta seguro de eliminar la formula de vacaciones :" & RSFORMS!NOMBRE, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    RSFORMS.Delete
    RSFORMS.Requery
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
    If RSFORMS.EOF Or RSFORMS.BOF Then Exit Sub
    Frame1.Tag = "EDITAR"
    Frame1.Enabled = True
    OCULTAR
    xNombre.SetFocus
End Sub

Private Sub CMNUEVO_CLICK()
    Frame1.Enabled = True
    Frame1.Tag = "NUEVO"
    Call Limpiar
    OCULTAR
    xNombre.SetFocus
End Sub
Private Sub Limpiar()
    xNombre.Text = ""
    xFormula.Text = ""
    xCriterio.Text = ""
    xTipo.ListIndex = 0
    Check1.Value = 0
    Check2.Value = 0
End Sub

Private Sub Command1_Click()
    frmHelpTmp.LlamaFrm = 1
    frmHelpTmp.Show 1
End Sub

Private Sub Form_Load()
    Set RSFORMS = New ADODB.Recordset
    RSFORMS.Open "SELECT *  FROM FORMULASVAC", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSFORMS
    xTipo.ListIndex = 0
    XDATA_ROWCOLCHANGE 0, 0
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

Private Sub xCriterio_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub

Private Sub XDATA_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    On Error Resume Next
    If RSFORMS.BOF Or RSFORMS.EOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If Frame1.Enabled Then Exit Sub
    xNombre.Text = RSFORMS!NOMBRE
    xFormula.Text = RSFORMS!FORMULA
    xTipo.ListIndex = RSFORMS!TIPO
    If Not IsNull(RSFORMS!CRITERIO) Then
        xCriterio.Text = RSFORMS!CRITERIO
    Else
        xCriterio.Text = " "
    End If
    Check1.Value = IIf(RSFORMS!AFECTOPRO, 1, 0)
    Check2.Value = IIf(RSFORMS!GENE, 1, 0)
End Sub

Private Sub XFormula_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub

Private Sub xNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = RestringeCaracter(KeyAscii, CGCADVAL)
End Sub
