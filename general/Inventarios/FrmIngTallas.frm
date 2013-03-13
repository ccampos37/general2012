VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmIngTallas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso por Tallas"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "FrmIngTallas.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Enviar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5220
      TabIndex        =   9
      Top             =   3180
      Width           =   1170
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   6570
      TabIndex        =   8
      Top             =   3180
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   0
      TabIndex        =   7
      Top             =   -105
      Width           =   7815
      Begin VB.Image Image1 
         Height          =   480
         Left            =   195
         Picture         =   "FrmIngTallas.frx":1272
         Top             =   195
         Width           =   480
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "UTILITARIO ESPECIAL PARA EL REGISTRO DE TALLAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   630
         Left            =   900
         TabIndex        =   10
         Top             =   270
         Width           =   4035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   2475
      Left            =   -15
      TabIndex        =   0
      Top             =   600
      Width           =   7830
      Begin MSDataGridLib.DataGrid DTGtallas 
         Height          =   720
         Left            =   150
         TabIndex        =   5
         Top             =   1470
         Width           =   7560
         _ExtentX        =   13335
         _ExtentY        =   1270
         _Version        =   393216
         Appearance      =   0
         BackColor       =   15925247
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdTallas 
         Caption         =   "&Mostrar Tallas"
         Enabled         =   0   'False
         Height          =   435
         Left            =   1665
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox TxArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00F2FFFF&
         Height          =   330
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   255
         Width           =   2040
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Relacion de Tallas"
         ForeColor       =   &H00E4FEFC&
         Height          =   270
         Left            =   195
         TabIndex        =   6
         Top             =   1215
         Width           =   1515
      End
      Begin VB.Label LbArticulo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   3720
         TabIndex        =   3
         Top             =   255
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo del Articulo :"
         ForeColor       =   &H00E4FEFC&
         Height          =   270
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1515
      End
   End
End
Attribute VB_Name = "FrmIngTallas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CnxAux As ADODB.Connection
Dim RSTALLAS As ADODB.Recordset

Private Sub GrabarformNI(ByRef Cont As Integer)
Dim FORMA As New FrmCreacionSin
Dim CODARTICULO As String
Dim I As Integer
    Cont = 0
    For I = 0 To RSTALLAS.Fields.count - 1
        If RSTALLAS.Fields(I).Value > 0 Then
            Cont = Cont + 1
            CODARTICULO = Trim(TxArticulo) & Trim(RSTALLAS.Fields(I).Name)
            FORMA.TxtArticulo = CODARTICULO
            FORMA.Label13.Caption = LbArticulo.Caption
            FORMA.TxtCantidad = Format(RSTALLAS.Fields(I).Value, "#0.00 ")
            FORMA.LblCantidad = Format(RSTALLAS.Fields(I).Value, "#0.00 ")
            VGabrev = Devolver_Dato(1, CODARTICULO, "MAEART", "ACODIGO", False, "AUNIDAD")
            FORMA.lbcantstk = Devolver_Dato(1, CODARTICULO, "STKART", "STCODIGO", False, "STSKDIS")
            Call FORMA.Command1_Click
       End If
    Next
End Sub
Private Sub GrabarFrmGuiaSal(ByRef Cont As Integer)
Dim FORMA As New FrmCreacionSal
Dim CODARTICULO As String
Dim I As Integer
    Cont = 0
    For I = 0 To RSTALLAS.Fields.count - 1
        If RSTALLAS.Fields(I).Value > 0 Then
            Cont = Cont + 1
            CODARTICULO = Trim(TxArticulo) & Trim(RSTALLAS.Fields(I).Name)
            FORMA.TxtArticulo = CODARTICULO
            FORMA.TxDescri = LbArticulo.Caption
            FORMA.TxtCantidad = Format(RSTALLAS.Fields(I).Value, "#0.00 ")
            'FORMA.LblCantidad = Format(RSTALLAS.Fields(i).Value, "#0.00 ")
            VGabrev = Devolver_Dato(1, CODARTICULO, "MAEART", "ACODIGO", False, "AUNIDAD")
            FORMA.lbcantstk = Devolver_Dato(1, CODARTICULO, "STKART", "STCODIGO", False, "STSKDIS")
            Call FORMA.Command1_Click
       End If
    Next
End Sub


Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdGrabar_Click()
Dim con As Integer
    Screen.MousePointer = 11
    Select Case VGRegEnt
        Case 0, 1: Call GrabarformNI(con)
        Case 2: Call GrabarFrmGuiaSal(con)
    End Select
    MsgBox "Se enviaron " & con & " registros"
    Screen.MousePointer = 1
    Call Limpiar
End Sub

Private Sub CmdTallas_Click()
    Call CrearTempoTallas(TxArticulo)
    Set RSTALLAS = New ADODB.Recordset
    RSTALLAS.Open "SELECT * FROM TMPTALLAS", CnxAux, adOpenKeyset, adLockOptimistic
    Set DTGtallas.DataSource = RSTALLAS
    DTGtallas.Refresh
    CmdGrabar.Enabled = True
End Sub

Private Sub DTGtallas_AfterColUpdate(ByVal ColIndex As Integer)
    If VGRegEnt <> 1 Then
        If Not Validar_Stock(ColIndex) Then Exit Sub
    End If
    If DTGtallas.Col + 1 <= DTGtallas.Columns.count - 1 Then DTGtallas.Col = DTGtallas.Col + 1
End Sub
Private Function Validar_Stock(Columna As Integer) As Boolean
Dim RSAUX As ADODB.Recordset
Dim STK As Long
On Error GoTo tiponocon
Validar_Stock = True
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT STSKDIS FROM STKART WHERE STCODIGO='" & _
               Trim(TxArticulo) & Trim(DTGtallas.Columns(Columna).DataField) & "' AND STALMA='" & VGAlma & "'", VGCNx, adOpenKeyset, adLockReadOnly
    STK = ESNULO(RSAUX!STSKDIS, 0)
    
    If DTGtallas.Columns(Columna).Value > STK Then
       MsgBox "El stock actual es: " & STK & "." & Chr(13) & Chr(13) & _
       "con la cantidad ingresada :" & DTGtallas.Columns(Columna).Value & Chr(13) & _
       "El Stock es negativo :(" & STK - DTGtallas.Columns(Columna).Value & ")", vbExclamation, "Desbordamiento de Stock"
       DTGtallas.Columns(Columna).Value = 0
       DTGtallas.Col = Columna
       Validar_Stock = False
       Exit Function
    End If
    Exit Function
tiponocon:
    Select Case Err.Number
        Case 13
            DTGtallas.Columns(Columna).Value = 0
            Exit Function
    End Select
End Function
Private Sub Form_Load()
    If VGRegEnt = 1 Then
        Me.Caption = "Ingreso por Tallas"
     Else
        Me.Caption = "Salida por Tallas"
    End If
    Set CnxAux = ConectarAux
End Sub

Private Sub TxArticulo_DblClick()
    Dim RSART As ADODB.Recordset
    Set RSART = New ADODB.Recordset
    RSART.Open " SELECT  DISTINCT Mid([ACODIGO],1,LEN(ACODIGO)-3) AS COD, MAEART.ADESCRI " & _
               " From MAEART WHERE ATALLA<>''", VGCNx, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSART
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxArticulo.text = vGUtil(1): LbArticulo.Caption = vGUtil(2)
        CmdTallas.Enabled = True
      Else
        TxArticulo.text = "": LbArticulo.Caption = ""
        CmdTallas.Enabled = False
    End If
    Set DTGtallas.DataSource = Nothing
    Set RSTALLAS = Nothing
    CmdGrabar.Enabled = False
End Sub

Private Sub TxArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 112 Then TxArticulo_DblClick
End Sub
Private Sub CrearTempoTallas(codigo As String)
Dim RSAUX As ADODB.Recordset
Dim sqlcad As String
    sqlcad = ""
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open " SELECT  Right(MAEART.ACODIGO,3) as tallas " & _
               " From MAEART Where " & _
               " Mid(MAEART.ACODIGO, 1, Len(MAEART.ACODIGO) - 3) = '" & Trim(codigo) & "'", VGCNx, adOpenKeyset, adLockReadOnly
    Do While Not RSAUX.EOF
        sqlcad = sqlcad & RSAUX.Fields(0) & " LONG,"
        RSAUX.MoveNext
    Loop
    sqlcad = Mid(sqlcad, 1, Len(sqlcad) - 1)
    If ExisteElem(0, CnxAux, "TMPTALLAS") Then CnxAux.Execute "DROP TABLE TMPTALLAS"
    Set RSAUX = New ADODB.Recordset
    CnxAux.Execute ("CREATE TABLE TMPTALLAS(" & sqlcad & ")")
    'Insertando un nuevo registro
    Dim I As Integer
    RSAUX.Open "TMPTALLAS", CnxAux, adOpenKeyset, adLockOptimistic
    RSAUX.AddNew
    For I = 0 To RSAUX.Fields.count - 1
        RSAUX.Fields(I) = 0
    Next
    RSAUX.Update
End Sub

Private Sub TxUnid_DblClick()
    Dim RSUNI As ADODB.Recordset
    Set RSUNI = New ADODB.Recordset
    RSUNI.Open "SELECT UM_ABREV,UM_NOMBRE FROM TABUNIMED", VGCNx, adOpenKeyset, adLockReadOnly
    frmref.Conectar RSUNI
    frmref.Show 1
    If vGUtil(1) <> "" Then
        TxUnid.text = vGUtil(1)
      Else
        TxUnid.text = ""
    End If
End Sub
Private Sub Limpiar()
    Set RSTALLAS = Nothing
    TxArticulo = ""
    LbArticulo = ""
    CmdTallas.Enabled = False
    CmdGrabar.Enabled = False
    Set DTGtallas.DataSource = Nothing
End Sub
