VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmVerlotes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lote Seleccionado"
   ClientHeight    =   4650
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   5220
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4176
      Width           =   1176
   End
   Begin VB.TextBox txtCol 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3852
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1872
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   36
      TabIndex        =   5
      Top             =   4176
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4032
      TabIndex        =   8
      Top             =   4176
      Width           =   1080
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   36
      TabIndex        =   0
      Top             =   36
      Width           =   5112
      Begin VB.TextBox cCod 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   16
         Top             =   216
         Width           =   2052
      End
      Begin VB.TextBox cDesc 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   15
         Top             =   540
         Width           =   3204
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3600
         MaxLength       =   13
         TabIndex        =   12
         Top             =   1404
         Width           =   864
      End
      Begin VB.TextBox ncant 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2304
         MaxLength       =   13
         TabIndex        =   2
         Top             =   972
         Width           =   936
      End
      Begin VB.TextBox Lote 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1260
         MaxLength       =   20
         TabIndex        =   1
         Top             =   1692
         Width           =   3204
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   216
         X2              =   4500
         Y1              =   900
         Y2              =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Articulo  :"
         Height          =   192
         Left            =   288
         TabIndex        =   14
         Top             =   612
         Width           =   636
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   192
         Left            =   288
         TabIndex        =   13
         Top             =   216
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Stock x  Lote"
         Height          =   192
         Left            =   2628
         TabIndex        =   11
         Top             =   1404
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad a Distribuir"
         Height          =   192
         Left            =   288
         TabIndex        =   10
         Top             =   1044
         Width           =   1440
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   192
         Left            =   324
         TabIndex        =   9
         Top             =   1764
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "&Adicionar"
      Height          =   375
      Left            =   1296
      TabIndex        =   6
      Top             =   4176
      Width           =   1176
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Gridlote 
      Height          =   2280
      Left            =   36
      TabIndex        =   4
      Top             =   1692
      Width           =   5124
      _ExtentX        =   9049
      _ExtentY        =   4022
      _Version        =   393216
      Rows            =   0
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   240
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "^Cód Art|Descripción|Uni.|>Cantidad|>Saldo|>Cant.Recibida"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Gridcabeza 
      Height          =   300
      Left            =   36
      TabIndex        =   17
      Top             =   1440
      Width           =   5124
      _ExtentX        =   9049
      _ExtentY        =   529
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      RowHeightMin    =   240
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "^Cód Art|Descripción|Uni.|>Cantidad|>Saldo|>Cant.Recibida"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
End
Attribute VB_Name = "frmVerlotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adodc1 As ADODB.Recordset
Attribute adodc1.VB_VarHelpID = -1
Public almacen As String

Private Sub Command3_Click()
If Gridlote.Rows <> 0 Then
   If Gridlote.Rows = 1 Then
      Gridlote.Clear
      Gridlote.Rows = 0
      Gridlote.Cols = 3
   Else
      Gridlote.RemoveItem (Gridlote.Row)
   End If
End If
End Sub

Private Sub cmdChange_Click()
Dim tot As Double
tot = 0
For n = 0 To Gridlote.Rows - 1
    tot = tot + CDbl(Val(Gridlote.TextMatrix(n, 1)))
Next

If tot >= ncant Then
  cmdChange.Enabled = False
  Exit Sub
Else
  cmdChange.Enabled = True
End If

    Me.Visible = False
    frmReglotes.Frame2.Visible = False
    frmReglotes.almacen = almacen
    frmReglotes.LoadLotesArti (cCod)
    frmReglotes.Text1 = cCod
    frmReglotes.Label3 = cDesc
    frmReglotes.Caption = "Seleccione o Adicione el Lote Destino "
    frmReglotes.cmdExitimport.Visible = True
    frmReglotes.cmdsubsalida.Visible = False
    frmReglotes.lblmsg.Visible = True
    frmReglotes.cmdretorna.Visible = True
    frmReglotes.Show 1
End Sub

Private Sub Command1_Click()
grabar

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
'Text1.SetFocus

End Sub

Private Sub Form_Load()
ncant = ""
Me.Top = frmTraIng.Top
Me.Left = frmTraIng.Left
Gridlote.Clear
Gridlote.Rows = 0
Gridlote.Cols = 3
Gridcabeza.Cols = 3
Gridcabeza.FormatString = "<Lote                             |>CantIng      |>Stock          "
Gridcabeza.ColWidth(0) = 2000
Gridcabeza.ColWidth(1) = 900
Gridcabeza.ColWidth(2) = 900
cmdChange.Enabled = True

Gridlote.ColAlignment(0) = 2
Gridlote.ColWidth(0) = 2000
Gridlote.ColWidth(1) = 900
Gridlote.ColWidth(2) = 900

End Sub

Private Sub Gridlote_KeyPress(KeyAscii As Integer)
Dim csql As String
    
    If Gridlote.Col = 1 Then
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
            txtCol.FontName = Gridlote.CellFontName
            txtCol.FontSize = Gridlote.CellFontSize
            txtCol.Width = Gridlote.CellWidth
            txtCol.Height = Gridlote.CellHeight
            txtCol.Left = Gridlote.Left + Gridlote.CellLeft
            txtCol.Top = Gridlote.Top + Gridlote.CellTop
            txtCol.Visible = True
            txtCol = Chr(KeyAscii)
            txtCol.SelStart = 1
            txtCol.SetFocus
        End If
    End If

End Sub

Private Sub ncant_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 4)) < Val(ncant) Then
            ncant = Format(ncant, "0.00")
            ncant.SelStart = 0
            ncant.SelLength = Len(ncant)
        Else
            ncant = Format("0" & ncant, "0.00")
        End If
    End If
End Sub
Sub grabar()
On Error GoTo Err
Dim criterio  As String
Dim tot As Double

If Trim(almacen) = "" Then
   MsgBox "Seleccione primero el Almacen, para poder registrar las Series  ", vbInformation, "Aviso...!"
   Exit Sub
End If
     tot = 0
     For n = 0 To Gridlote.Rows - 1
         tot = tot + CDbl(Gridlote.TextMatrix(n, 1))
     Next

    If tot = 0 Or tot > ncant Then
       MsgBox "Distribuya de manera correcta los Ingresos", vbInformation, "Aviso...!"
       Exit Sub
    End If
    
  Set adodc1 = New ADODB.Recordset
  With adodc1
     'RMM***********************************************************************08/08/2001
     'Vgcnx.Execute "delete from art_serie"
     VGcnx.Execute "delete from art_LOTE where ALMA='" & almacen & "' AND  acodigo='" & VGcod & "'"
     '***********************************************************************
     criterio = "select * from art_LOTE "
     adodc1.Open criterio, VGcnx, adOpenDynamic, adLockBatchOptimistic
     For n = 0 To Gridlote.Rows - 1
         If CDbl(Gridlote.TextMatrix(n, 1)) > 0 Then
           .AddNew
           .Fields("alma") = almacen
           .Fields("acodigo") = VGcod
           .Fields("LOTE") = Gridlote.TextMatrix(n, 0)
           .Fields("CANTID") = CDbl(Gridlote.TextMatrix(n, 1))
           .UpdateBatch
         End If
     Next
  End With
  
frmTraIng.Flex1.TextMatrix(frmTraIng.Flex1.Row, 5) = Format(tot, "###,##0.00")
Unload Me
  
Exit Sub
Err:
  MsgBox Err.Description, vbInformation, "Aviso"
  
End Sub


Private Sub txtCol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(ncant) < Val(txtCol) Then
            txtCol = Format(txtCol, "0.00")
            txtCol.SelStart = 0
            txtCol.SelLength = Len(txtCol)
        Else
            Gridlote.text = Format("0" & txtCol, "0.00")
            txtCol.Visible = False
            Gridlote.SetFocus
        End If
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        txtCol.Visible = False
        Gridlote.SetFocus
    Else
        'Reales_Positivos KeyAscii, txtCol
    End If
End Sub

Private Sub txtCol_LostFocus()
txtCol.Visible = False
End Sub
