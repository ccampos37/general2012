VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frm_manten_ordfabri 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Ordenes de Fabricación"
   ClientHeight    =   4860
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6060
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4872
      Left            =   36
      TabIndex        =   8
      Top             =   36000
      Width           =   6324
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   528
         Left            =   1560
         TabIndex        =   24
         Top             =   4290
         Width           =   1032
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   528
         Left            =   5280
         TabIndex        =   10
         Top             =   4296
         Width           =   975
      End
      Begin VB.CommandButton cmdedit 
         Caption         =   "&Modificar"
         Height          =   528
         Left            =   2604
         TabIndex        =   9
         Top             =   4308
         Width           =   1032
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridArti 
         Height          =   4116
         Left            =   1548
         TabIndex        =   20
         Top             =   144
         Width           =   4704
         _ExtentX        =   8308
         _ExtentY        =   7250
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
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
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
      Begin VB.Image Image1 
         Height          =   4668
         Left            =   -612
         Picture         =   "frm_manten_ordfabri.frx":0000
         Stretch         =   -1  'True
         Top             =   144
         Width           =   2436
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4872
      Left            =   24
      TabIndex        =   0
      Top             =   -36
      Width           =   6012
      Begin VB.CommandButton cmdExitimport 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   4884
         TabIndex        =   22
         Top             =   4320
         Visible         =   0   'False
         Width           =   924
      End
      Begin VB.CommandButton cmdedita 
         Caption         =   "&Editar"
         Height          =   495
         Left            =   1056
         TabIndex        =   19
         Top             =   4320
         Width           =   948
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3216
         Left            =   96
         TabIndex        =   11
         Top             =   1020
         Width           =   5724
         _ExtentX        =   10107
         _ExtentY        =   5662
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   864
         TabCaption(0)   =   "Ordenes de Fabricación"
         TabPicture(0)   =   "frm_manten_ordfabri.frx":2DC3
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblmsg"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "grid"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Datos del Documento"
         TabPicture(1)   =   "frm_manten_ordfabri.frx":2DDF
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "Label4"
         Tab(1).Control(2)=   "Label18"
         Tab(1).Control(3)=   "Label5"
         Tab(1).Control(4)=   "lbnomclie"
         Tab(1).Control(5)=   "dFech_ven"
         Tab(1).Control(6)=   "dFech_fab"
         Tab(1).Control(7)=   "Text3"
         Tab(1).Control(8)=   "Text2"
         Tab(1).ControlCount=   9
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   -73716
            MaxLength       =   20
            TabIndex        =   25
            Top             =   1212
            Width           =   1275
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   -73716
            MaxLength       =   20
            TabIndex        =   12
            Top             =   816
            Width           =   1275
         End
         Begin MSMask.MaskEdBox dFech_fab 
            Height          =   288
            Left            =   -73728
            TabIndex        =   14
            Top             =   2016
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   -2147483634
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox dFech_ven 
            Height          =   288
            Left            =   -73716
            TabIndex        =   13
            Top             =   1608
            Width           =   1248
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   -2147483634
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "dd/mm/yyyy"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grid 
            Height          =   2280
            Left            =   120
            TabIndex        =   21
            Top             =   600
            Width           =   5412
            _ExtentX        =   9551
            _ExtentY        =   4022
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
         Begin VB.Label lbnomclie 
            Caption         =   "????"
            Height          =   252
            Left            =   -72420
            TabIndex        =   27
            Top             =   1272
            Width           =   3048
         End
         Begin VB.Label Label5 
            Caption         =   "Cod.Cliente"
            Height          =   252
            Left            =   -74748
            TabIndex        =   26
            Top             =   1212
            Width           =   1092
         End
         Begin VB.Label lblmsg 
            Caption         =   "Seleccione los Lotes haciendo doble Click, y de <Retornar>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   252
            Left            =   144
            TabIndex        =   23
            Top             =   2928
            Visible         =   0   'False
            Width           =   5436
         End
         Begin VB.Label Label18 
            Caption         =   "Nro.Ord Fab"
            Height          =   252
            Left            =   -74748
            TabIndex        =   18
            Top             =   816
            Width           =   1092
         End
         Begin VB.Label Label4 
            Caption         =   "Fecha Inicio"
            Height          =   252
            Left            =   -74748
            TabIndex        =   16
            Top             =   1656
            Width           =   1212
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Final"
            Height          =   252
            Left            =   -74760
            TabIndex        =   15
            Top             =   2088
            Width           =   1572
         End
      End
      Begin VB.CommandButton cmdGrabalote 
         Caption         =   "Grabar"
         Height          =   495
         Left            =   2028
         TabIndex        =   17
         Top             =   4320
         Width           =   972
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Nuevo"
         Height          =   495
         Left            =   96
         TabIndex        =   7
         Top             =   4320
         Width           =   924
      End
      Begin VB.CommandButton CMDBORRA 
         Caption         =   "&Eliminar"
         Height          =   495
         Left            =   3804
         TabIndex        =   6
         Top             =   4320
         Width           =   948
      End
      Begin VB.CommandButton cmdsubsalida 
         Caption         =   "&Salir"
         Height          =   495
         Left            =   4884
         TabIndex        =   5
         Top             =   4320
         Width           =   924
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1368
         MaxLength       =   20
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   216
         Width           =   2268
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   252
         X2              =   5928
         Y1              =   972
         Y2              =   972
      End
      Begin VB.Label Label6 
         Caption         =   "Código"
         Height          =   252
         Left            =   216
         TabIndex        =   4
         Top             =   252
         Width           =   732
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label3"
         Enabled         =   0   'False
         Height          =   288
         Left            =   1368
         TabIndex        =   3
         Top             =   576
         Width           =   4416
      End
      Begin VB.Label Label2 
         Caption         =   "Descripción"
         Height          =   252
         Left            =   216
         TabIndex        =   2
         Top             =   612
         Width           =   972
      End
   End
End
Attribute VB_Name = "frm_manten_ordfabri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public almacen As String
Private Sub cmdBorra_Click()
If grid.Rows = 1 Then Exit Sub
If ClsTock.ArticuloConOrdFabri(VGAlma, Text1, grid.TextMatrix(grid.Row, 0), VGCNx) Then
   MsgBox "El Articulo ya Tiene Movimientos Registrados con la respectiva Ordenes de Fabricación", vbCritical, "Error en Datos.."
   Exit Sub
End If

If MsgBox("Desea Eliminar el Lote '" & grid.TextMatrix(grid.Row, 0) & "'", vbInformation + vbYesNo, "Eliminar Lotes ") = vbYes Then
   VGCNx.Execute "DELETE from ORDE_FAB where ORD_ALMA='" & almacen & "' and ORD_CODART='" & Text1.text & "' and ORD_FABNUM='" & grid.TextMatrix(grid.Row, 0) & "'"
   LoadOrdenFab (Text1)
End If
End Sub

Private Sub cmdedit_Click()
Call ClearForm
Text1.Enabled = False
Label3.Enabled = False
Text1 = GridArti.TextMatrix(GridArti.Row, 0)
Label3 = GridArti.TextMatrix(GridArti.Row, 1)
LoadOrdenFab (GridArti.TextMatrix(GridArti.Row, 0))
Frame2.Visible = False
SSTab1.Tab = 0
End Sub

Private Sub cmdNew_Click()
Call ClearForm
Frame2.Visible = False
End Sub

Private Sub cmdModifica_Click()

End Sub

Private Sub cmdadd_Click()
SSTab1.Tab = 1
ClearForm
Text3.Enabled = True
Text3.SetFocus
End Sub

Private Sub cmdedita_Click()
   Call CARGADATOS(Text1, grid.TextMatrix(grid.Row, 0))
   SSTab1.Tab = 1
   Text3.Enabled = False
End Sub

Private Sub cmdExitimport_Click()
If SSTab1.Tab = 0 Then
   Unload Me
Else
   SSTab1.Tab = 0
End If
End Sub

Private Sub cmdGrabalote_Click()
Call grabalote(Text1)
SSTab1.Tab = 0
ClearForm
End Sub

Private Sub cmdNuevo_Click()
Call ClearForm
Text1.Enabled = True
Label3.Enabled = True
Frame2.Visible = False
Text1 = ""
Label3 = ""
End Sub

Private Sub cmdsubsalida_Click()
Call ClearForm
If SSTab1.Tab = 1 Then
   SSTab1.Tab = 0
Else
   Frame2.Visible = True
End If
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command6_Click()

End Sub

Private Sub dFech_fab_KeyPress(KeyAscii As Integer)
        Tabula (KeyAscii)
End Sub

Private Sub dFech_ven_KeyPress(KeyAscii As Integer)
        Tabula (KeyAscii)
End Sub

Private Sub Form_Load()
   Screen.MousePointer = 11
   almacen = VGAlma
   Text1.text = ""
   central Me
   grid.Cols = 2
   Call LoadGrid
   SSTab1.Tab = 0
   cmdExitimport.Visible = False
   cmdsubsalida.Visible = True
   Me.Caption = "Registrar Ordenes Fabricación"
   Screen.MousePointer = 1
End Sub

Sub LoadGrid()
Dim rs As New ADODB.Recordset
Dim SQL As String
'ORDE_FAB.ORD_FABNUM,
SQL = "SELECT ORDE_FAB.ORD_CODART, MAEART.ADESCRI FROM MAEART INNER JOIN ORDE_FAB ON MAEART.ACODIGO = ORDE_FAB.ORD_CODART WHERE ORD_ALMA='" & VGAlma & "'" '"SELECT MAEART.ACODIGO, MAEART.ADESCRI, format(STKART.STSKDIS,'###,##0.00') FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO WHERE STKART.STALMA='" & almacen & "' AND MAEART.AFLOTE='S'"
rs.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
Set GridArti.DataSource = rs
GridArti.FormatString = "<Cod.Articulo        |<Descripción                                                           "
rs.Close
End Sub

Sub ClearForm()
Text3 = "": Text2 = "": lbnomclie = "": dFech_ven = "__/__/____": dFech_fab = "__/__/____": txobse = ""
If SSTab1.Tab = 1 Then
   cmdGrabalote.Enabled = True
Else
   cmdGrabalote.Enabled = False
End If
End Sub

Public Sub LoadOrdenFab(ByVal arArti As String)
Dim rs As New ADODB.Recordset
Dim SQL As String
SQL = "SELECT ORD_FABNUM,FECH_INI,FECH_FIN FROM ORDE_FAB WHERE ORD_ALMA='" & almacen & "'  AND ORD_CODART='" & arArti & "'"
rs.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
grid.Clear
grid.Rows = 2
If Not rs.EOF Then Set grid.DataSource = rs
grid.FormatString = "<Orden de Fabricación                        |^Fecha de Inicio |^Fecha  Final   "
rs.Close
End Sub
'grabalote
Sub grabalote(ByVal arCod As String)
Dim SQL As String
If Trim(Text3) = "" Then Exit Sub
If ClsTock.ExisteOrdenFab(almacen, arCod, Text3, VGCNx) Then
   SQL = "Update ORDE_FAB set ORD_CODCLIE='" & Text2 & "',FECH_INI=" & DateSQL2000(dFech_ven) & ",FECH_FIN=" & DateSQL2000(dFech_fab) & ",CODI_OPER='" & VGUsuario & "',FECH_TRAN=FORMAT(NOW,'DD/MM/YYYY') where ORD_ALMA='" & almacen & "' and ORD_CODART='" & arCod & "' and ORD_FABNUM='" & Text3 & "'"
Else
   SQL = "Insert Into ORDE_FAB (ORD_ALMA,ORDE_FAB,ORD_CODART,ORD_CODCLIE,FECH_INI,FECH_FIN,FECH_TRAN,CODI_OPER) VALUES ('" & almacen & "','" & Text2 & "','" & arCod & "','" & Text3 & "'," & DateSQL2000(dFech_ven) & "," & DateSQL2000(dFech_fab) & ",FORMAT(NOW,'DD/MM/YYYY'),'" & VGUsuario & "')"
End If
   
   VGCNx.Execute SQL
   
     'If frmTraIng.Visible = True Then
     '   If Not ExisteEnGrid(Text3) Then
     '      frmVerlotes.Gridlote.AddItem Text3.text & Chr(9) & "0" & Chr(9) & "0.0"
     '   End If
     '   Unload Me
     '   frmVerlotes.show 1
     'Else
        LoadOrdenFab (Text1)
     'End If
End Sub

Private Sub grid_dblClick()
If grid.CellBackColor = &H80000005 And grid.TextMatrix(grid.Row, 0) <> "" Then
   For n = 0 To grid.Cols - 1
   grid.Col = n
   grid.CellBackColor = &H8000000D 'azul
   grid.CellForeColor = &H80000005 'blanco
   Next
Else
   For n = 0 To grid.Cols - 1
   grid.Col = n
   grid.CellBackColor = &H80000005 'blanco
   grid.CellForeColor = &H0& 'negro
   Next
End If
''''     If frmTraIng.Visible = True Then
''''        If Not ExisteEnGrid(grid.TextMatrix(frmReglotes.grid.Row, 0)) Then
''''           If grid.TextMatrix(frmReglotes.grid.Row, 0) = "" Then Exit Sub
''''           frmVerlotes.Gridlote.AddItem frmReglotes.grid.TextMatrix(frmReglotes.grid.Row, 0) & Chr(9) & "0" & Chr(9) & Format(Val(frmReglotes.grid.TextMatrix(frmReglotes.grid.Row, 1)), "###,##0.00")
''''        Else
''''           MsgBox "El Lote ya fue Seleccionado......!", vbInformation, "Error en Datos"
''''           Exit Sub
''''        End If
''''        'frmVerlotes.Lote = frmReglotes.grid.TextMatrix(frmReglotes.grid.Row, 0)
''''        'frmVerlotes.Text2 = Val(frmReglotes.grid.TextMatrix(frmReglotes.grid.Row, 1)) 'stock
''''        Unload Me
''''        frmVerlotes.show 1
''''        'frmVerlotes.Gridlote.SetFocus
''''     End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   grid_dblClick
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If PreviousTab = 0 Then
   Call CARGADATOS(Text1, grid.TextMatrix(grid.Row, 0))
End If
End Sub
Sub CARGADATOS(ByVal arArti As String, ByVal ORDEN As String)
Call ClearForm
Dim SQL As String
Dim rs As New ADODB.Recordset
SQL = "SELECT * FROM ORDE_FAB WHERE ORD_ALMA='" & almacen & "'  AND ORD_CODART='" & arArti & "' AND ORD_FABNUM='" & ORDEN & "'"
rs.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
If Not rs.EOF Then
   Text3 = rs!ORDE_FABNUM
   dFech_ven = FechMask(rs!FECH_INI)
   dFech_fab = FechMask(rs!FECH_FIN)
   Text2 = cNull(rs!ORD_CODCLIE)
End If
rs.Close
End Sub
''Function ExisteEnGrid(ByVal Lote As String) As Boolean
''ExisteEnGrid = False
''For n = 0 To frmVerlotes.Gridlote.Rows - 1
''    If frmVerlotes.Gridlote.TextMatrix(n, 0) = Lote Then
''       ExisteEnGrid = True
''    End If
''Next
''End Function

Private Sub Text3_KeyPress(KeyAscii As Integer)
        Tabula (KeyAscii)
End Sub

Private Sub txobse_KeyPress(KeyAscii As Integer)
        Tabula (KeyAscii)
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 46 Then
    Label3 = ""
 ElseIf KeyCode = 112 Then
      Text1_DblClick
 End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 And Text1 <> "" Then
         criterio = "ACODIGO = " & "'" + Text1.text + "'"
         Data1.Recordset.FindFirst criterio
         If Not Data1.Recordset.NoMatch Then
            Label3.Caption = Data1.Recordset.Fields("ADESCRI")
            Text2.SetFocus
         Else
            MsgBox "El codigo no existe", vbOKOnly, "No Encontrado"
            Text1.SetFocus
         End If
        
     End If
End Sub

Private Sub Text1_DblClick()
   VGForm1 = 21
   FormAyuArt1.Show 1
'   If Text1 <> "" And List1.Rows = 0 Then
'       agregarlista
'   End If
End Sub
