VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAyuda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5088
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5448
   Icon            =   "FrmAyudaRecordSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5088
   ScaleWidth      =   5448
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   90
      TabIndex        =   12
      Top             =   0
      Width           =   5235
      Begin VB.Label Label3 
         Caption         =   "Nota.- Al poner porcentaje=""%"" en cuadro de texto se mostraran todos los datos en la Grilla."
         Height          =   405
         Left            =   1080
         TabIndex        =   13
         Top             =   285
         Width           =   3900
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   168
         Picture         =   "FrmAyudaRecordSet.frx":0442
         Top             =   216
         Width           =   384
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "=>"
      Height          =   285
      Left            =   3120
      TabIndex        =   10
      Top             =   1125
      Visible         =   0   'False
      Width           =   330
   End
   Begin MSComCtl2.DTPicker DTPFecha 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Top             =   1170
      Width           =   3330
      _ExtentX        =   5884
      _ExtentY        =   508
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   8388608
      CalendarTitleForeColor=   16777215
      CheckBox        =   -1  'True
      Format          =   19791873
      CurrentDate     =   37474
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   1140
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   508
      _Version        =   393216
      Style           =   2
      ListField       =   "descripcion1"
      BoundColumn     =   "campo"
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   1275
      TabIndex        =   4
      Top             =   4710
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   4710
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DTG_detalle 
      Height          =   3150
      Left            =   105
      TabIndex        =   2
      Top             =   1515
      Width           =   5220
      _ExtentX        =   9208
      _ExtentY        =   5546
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   1
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxValor 
      BackColor       =   &H00E2FDFE&
      Height          =   300
      Left            =   105
      TabIndex        =   0
      Top             =   1155
      Width           =   3360
   End
   Begin VB.Frame FramOpciones 
      Height          =   435
      Left            =   105
      TabIndex        =   6
      Top             =   1050
      Width           =   3360
      Begin VB.OptionButton OptFalso 
         Caption         =   "Falso"
         Height          =   195
         Left            =   1995
         TabIndex        =   8
         Top             =   165
         Width           =   1050
      End
      Begin VB.OptionButton OptVerdadero 
         Caption         =   "Verdadero"
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   165
         Width           =   1035
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Nro. Reg. :"
      Height          =   210
      Left            =   3360
      TabIndex        =   15
      Top             =   4755
      Width           =   810
   End
   Begin VB.Label lab_reg 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 "
      Height          =   255
      Left            =   4185
      TabIndex        =   14
      Top             =   4710
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   885
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Criterio"
      Height          =   195
      Left            =   3495
      TabIndex        =   1
      Top             =   885
      Width           =   765
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rsconsulta As ADODB.Recordset
Attribute rsconsulta.VB_VarHelpID = -1
Dim tipocontrol As Integer
Dim valoropt As Variant
Private Sub CmdAceptar_Click()
    Set m_fileld = rsconsulta.Fields
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Set m_fileld = Nothing
    Unload Me
End Sub

Private Sub Command3_Click()
Dim criterio As String
    criterio = ""
    Select Case tipocontrol
        Case 1 'texto
            If Trim(TxValor) <> "" Then
                criterio = DataCombo1.BoundText & " like '" & Trim(TxValor) & "%'"
            End If
        Case 2 'booleano
            If Not Trim(valoropt) = "" Then criterio = DataCombo1.BoundText & " in(" & Trim(valoropt) & ")"
        Case 3 'fecha
            If Not IsNull(DTPFecha) Then criterio = DataCombo1.BoundText & "='" & Format(DTPFecha, "dd/mm/yyyy") & "'"
    End Select
    Call ejecutarconsulta(criterio)
    DTG_detalle.SetFocus
End Sub
Private Sub DataCombo1_Click(Area As Integer)
    Set rs = rscampos.Clone(adLockReadOnly)
    rs.Filter = "campo='" & Trim(DataCombo1.BoundText) & "'"
    If rs.RecordCount = 0 Then
        Call PrenderInputData(0)
      Else
        Call PrenderInputData(rs!tipocontrol)
        tipocontrol = rs!tipocontrol
    End If
    Set rs = Nothing
End Sub

Private Sub DTG_detalle_DblClick()
    Call DTG_detalle_KeyDown(13, 0)
End Sub

Private Sub DTG_detalle_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    If rsconsulta.Sort = "" Then
        rsconsulta.Sort = DTG_detalle.Columns.Item(ColIndex).DataField & " asc"
     ElseIf Right(rsconsulta.Sort, 3) = "asc" Then
        rsconsulta.Sort = DTG_detalle.Columns.Item(ColIndex).DataField & " desc"
     ElseIf Right(rsconsulta.Sort, 4) = "desc" Then
        rsconsulta.Sort = DTG_detalle.Columns.Item(ColIndex).DataField & " asc"
    End If
End Sub

Private Sub DTG_detalle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call CmdAceptar_Click
    End If
End Sub

Private Sub DTPFecha_Change()
    Call Command3_Click
End Sub

Private Sub DTPFecha_CloseUp()
    'Call Command3_Click
End Sub

Private Sub Form_Load()
    Set m_fileld = Nothing
    Screen.MousePointer = vbHourglass
    Me.Caption = m_TituloAyuda
    Set DataCombo1.RowSource = rscampos
    DataCombo1.BoundText = m_primercampo
    Call DataCombo1_Click(0)
    Call CrearCabecerdeGrid
    Screen.MousePointer = vbDefault
End Sub

Private Sub ejecutarconsulta(criterio As String)
On Error GoTo ErrorTipeo
    If Trim(TxValor.Text) = "%" Then
        m_RecordSet.Filter = 0
      Else
        m_RecordSet.Filter = criterio
    End If
    
'    Set rsconsulta = New ADODB.RecordSet
'    Set rsconsulta = m_RecordSet
'    Set DTG_detalle.DataSource = rsconsulta
    'Call RenombrarCabeceras
    lab_reg.Caption = Format(rsconsulta.RecordCount, "0 ")
    m_registros = rsconsulta.RecordCount
    Exit Sub
ErrorTipeo:
    m_RecordSet.Filter = DataCombo1.BoundText & "-1"
End Sub
Private Sub CrearCabecerdeGrid()
Dim i As Integer
    Set rsconsulta = New ADODB.Recordset
    Set rsconsulta = m_RecordSet
    Set DTG_detalle.DataSource = rsconsulta
    Call RenombrarCabeceras
End Sub
Private Sub RenombrarCabeceras()
    rscampos.MoveFirst
    For i = 0 To rscampos.RecordCount - 1
        DTG_detalle.Columns(i).Caption = rscampos!descripcion1
        rscampos.MoveNext
    Next
End Sub

Private Sub RenombrarCabecera()
Dim i As Integer
    For i = 1 To DTG_detalle.Columns.Count - 1
        Call DTG_detalle.Columns.Remove(i)
        If DTG_detalle.Columns.Count = 1 Then Exit For
    Next
End Sub
Private Sub PrenderInputData(Tipo As Integer)
    Select Case Tipo
        Case 0 'Ninguno
            TxValor.Visible = False
            FramOpciones.Visible = False
            DTPFecha.Visible = False
        Case 1 'Texto
            TxValor.Visible = True
            FramOpciones.Visible = False
            DTPFecha.Visible = False
        Case 2 'Booleano
            FramOpciones.Visible = True
            TxValor.Visible = False
            DTPFecha.Visible = False
        Case 3 'Fecha
            FramOpciones.Visible = False
            TxValor.Visible = False
            DTPFecha.Visible = True
            DTPFecha.Value = Null
    End Select
End Sub

Private Sub OptFalso_Click()
    valoropt = "0"
End Sub

Private Sub OptVerdadero_Click()
    valoropt = "1"
End Sub
Private Sub TxValor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call Command3_Click
    End If
End Sub

Private Sub TxValor_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    Exit Sub
    KeyAscii = KeyAscii
End Sub
