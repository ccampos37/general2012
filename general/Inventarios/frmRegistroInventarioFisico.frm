VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmRegistroInventarioFisico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Inventario Fisico"
   ClientHeight    =   6735
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9255
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   1164
      Left            =   2556
      TabIndex        =   30
      Top             =   2736
      Visible         =   0   'False
      Width           =   4512
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   228
         Left            =   324
         TabIndex        =   31
         Top             =   720
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbmsg 
         Caption         =   "Procesando...."
         Height          =   264
         Left            =   324
         TabIndex        =   32
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5736
      Left            =   72
      TabIndex        =   28
      Top             =   972
      Width           =   9084
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Gridview 
         Height          =   5088
         Left            =   4032
         TabIndex        =   29
         Top             =   360
         Width           =   4764
         _ExtentX        =   8387
         _ExtentY        =   8969
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
         ScrollBars      =   2
         SelectionMode   =   1
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
         _Band(0).Cols   =   2
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.Image imgLogo 
         Height          =   5052
         Left            =   324
         Picture         =   "frmRegistroInventarioFisico.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3180
      End
   End
   Begin VB.ComboBox CBORDEN 
      Height          =   288
      ItemData        =   "frmRegistroInventarioFisico.frx":2F41
      Left            =   7056
      List            =   "frmRegistroInventarioFisico.frx":2F4B
      TabIndex        =   8
      Text            =   "Por Descripción"
      Top             =   2592
      Width           =   2100
   End
   Begin VB.TextBox txing 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   7272
      TabIndex        =   25
      Top             =   3096
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   1908
      TabIndex        =   7
      Top             =   2592
      Width           =   3396
   End
   Begin VB.Frame Frame4 
      Height          =   912
      Left            =   108
      TabIndex        =   17
      Top             =   0
      Width           =   9072
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   636
         Left            =   3420
         Picture         =   "frmRegistroInventarioFisico.frx":2F6C
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   180
         Width           =   775
      End
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "frmRegistroInventarioFisico.frx":33AE
         Left            =   5976
         List            =   "frmRegistroInventarioFisico.frx":33B0
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   324
         Width           =   2976
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   636
         Left            =   1764
         Picture         =   "frmRegistroInventarioFisico.frx":33B2
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   775
      End
      Begin VB.CommandButton cmdBorra 
         Caption         =   "&Eliminar"
         Height          =   636
         Left            =   2592
         Picture         =   "frmRegistroInventarioFisico.frx":37F4
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   180
         Width           =   792
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   636
         Left            =   4248
         Picture         =   "frmRegistroInventarioFisico.frx":3C36
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   180
         Width           =   780
      End
      Begin VB.CommandButton cmdNew 
         Caption         =   "Nuevo"
         Height          =   636
         Left            =   108
         Picture         =   "frmRegistroInventarioFisico.frx":4078
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   180
         Width           =   780
      End
      Begin VB.CommandButton cmdModif 
         Caption         =   "&Modificar"
         Height          =   636
         Left            =   936
         Picture         =   "frmRegistroInventarioFisico.frx":44BA
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   792
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   5220
         TabIndex        =   23
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterio de Ingreso"
      Enabled         =   0   'False
      Height          =   1560
      Left            =   5400
      TabIndex        =   10
      Top             =   972
      Width           =   3756
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2088
         MaxLength       =   20
         TabIndex        =   6
         Top             =   972
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2088
         MaxLength       =   20
         TabIndex        =   5
         Top             =   396
         Width           =   1470
      End
      Begin VB.OptionButton OptFami 
         Caption         =   "Por Familia"
         Height          =   192
         Left            =   252
         TabIndex        =   12
         Top             =   1008
         Width           =   1488
      End
      Begin VB.OptionButton OptArti 
         Caption         =   "Por Articulo"
         Height          =   192
         Left            =   288
         TabIndex        =   11
         Top             =   396
         Value           =   -1  'True
         Width           =   1488
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   252
         Left            =   2124
         TabIndex        =   14
         Top             =   756
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Inicio "
         Height          =   252
         Left            =   2088
         TabIndex        =   13
         Top             =   180
         Width           =   732
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Documento"
      Enabled         =   0   'False
      Height          =   1560
      Left            =   72
      TabIndex        =   9
      Top             =   972
      Width           =   5232
      Begin VB.TextBox cObs 
         Height          =   540
         Left            =   144
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   936
         Width           =   4908
      End
      Begin VB.TextBox cNumInve 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1332
         TabIndex        =   2
         Top             =   324
         Width           =   1236
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   3564
         TabIndex        =   3
         Top             =   324
         Width           =   1488
         _ExtentX        =   2619
         _ExtentY        =   529
         _Version        =   393216
         Format          =   40239105
         CurrentDate     =   37090
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones :"
         Height          =   192
         Left            =   144
         TabIndex        =   27
         Top             =   684
         Width           =   1272
      End
      Begin VB.Label Label4 
         Caption         =   "Nro. Inv. Fisico"
         Height          =   192
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   192
         Left            =   2844
         TabIndex        =   15
         Top             =   360
         Width           =   696
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3720
      Left            =   72
      TabIndex        =   0
      Top             =   2952
      Width           =   9084
      _ExtentX        =   16007
      _ExtentY        =   6562
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label5 
      Caption         =   "Orden :"
      Height          =   192
      Left            =   5868
      TabIndex        =   26
      Top             =   2664
      Width           =   984
   End
   Begin VB.Label Label7 
      Caption         =   "Buscar  :"
      Height          =   228
      Left            =   108
      TabIndex        =   24
      Top             =   2628
      Width           =   1704
   End
End
Attribute VB_Name = "frmRegistroInventarioFisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim almacen As String
Dim rs As ADODB.Recordset
Dim Rs2 As ADODB.Recordset
Dim rSCod As ADODB.Recordset
Dim rsInve As New ADODB.Recordset
Dim cRt As String
Dim sCad As String
Dim Sproceso As String
Dim flagModif As Boolean 'Indicador si se ha modificado

'*********************************
'1- Guardado
'2- Nuevo
'*********************************

Private Sub CBORDEN_Click()
Text4_KeyPress (13)
End Sub

Private Sub cmdBorra_Click()
   If Gridview.Rows = 2 And Gridview.TextMatrix(Gridview.Row, 0) = "" Then Exit Sub
   If MsgBox("Esta Seguro que desea Eliminar el Documento Referente al Inventario Fisico Nro. " & Gridview.TextMatrix(Gridview.Row, 0), vbInformation + vbYesNo, "Eliminar Documento") = vbYes Then
      Call EliminaDoc(Gridview.TextMatrix(Gridview.Row, 0))
      CargaVista
   End If
End Sub

Private Sub CmdGrabar_Click()
Frame5.Visible = True
 
 Sproceso = "Guardado"
 Call SaveInvFisico
 If flagModif = False Then Call DesHabilita
 flagModif = False
 
 
 CargaVista
Frame5.Visible = False
End Sub

Private Sub cmdModif_Click()
     If Gridview.Rows = 2 And Gridview.TextMatrix(Gridview.Row, 0) = "" Then Exit Sub
     Call LoadCabezera(Gridview.TextMatrix(Gridview.Row, 0))
     Call LoadGrid("", Gridview.TextMatrix(Gridview.Row, 0))
     Call Habilita
     flagModif = False
     Sproceso = "Modifica"
End Sub

Private Sub cmdNew_Click()
     Frame5.Visible = True
     Sproceso = "Nuevo"
     flagModif = False
     If NuevoInventa = 1 Then
        Call Habilita
     End If
     'Sproceso = ""
     Frame5.Visible = False
End Sub

Private Sub cmdPrint_Click()
Dim CTIME As String
Dim ccadena As String
Dim aparam(2) As Variant
Dim aform(1) As Variant
Dim Reporte As String
Dim titulos As String

CTIME = Format(Time, "hh:mm:ss")

If OptArti.Value = True Then
   titulos = "Inv506 -- Control de Inventarios"
   Reporte = "Inv506.rpt"
Else
   titulos = "Inv507 -- Control de Inventarios"
   Reporte = "Inv507.rpt"
End If

If OptArti.Value = True Then
   If Text1 = "" Or Text2 = "" Then
      ccadena = " b.AUXNUMINVE=''" & cNumInve & "'' AND b.AUXALMA=''" & almacen & "''"
   Else
      ccadena = " b.AUXNUMINVE}=''" & cNumInve & "'' AND b.AUXALMA=''" & almacen & "'' AND b.AUXALMA=''" & almacen & "'' AND b.AUXCODART>=''" & Text1 & "'' and b.AUXCODART<=''" & Text2 & "''"
   End If
Else
   If Text1 = "" Or Text2 = "" Then
      ccadena = " b.AUXNUMINVE=''" & cNumInve & "'' AND b.AUXALMA=''" & almacen & "''"
   Else
      ccadena = " b.AUXNUMINVE=''" & cNumInve & "'' AND b.AUXALMA=''" & almacen & "'' AND b.AUXFAMIL>=''" & Text1 & "'' and b.AUXFAMIL<=''" & Text2 & "''"
   End If
End If
aform(0) = "Hora ='" & CTIME & "'"
aparam(0) = VGCNx.DefaultDatabase
aparam(1) = "" & ccadena & ""
Call ImpresionRptProc(Reporte, aform, aparam, , titulos)
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
CargaVista
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Command7_Click()
Dim Resul As Long
        If Sproceso = "Nuevo" Or Sproceso = "Modifica" Then
           Resul = MsgBox("Usted No Guardo los Datos...Desea Guardarlos Ahora", vbInformation + vbYesNoCancel, "Grabar Datos")
           If Resul = vbYes Then
               Call SaveInvFisico
               Sproceso = ""
               Gridview.Visible = True
               Call DesHabilita
           Else
               If Resul = vbNo Then
                  If Sproceso = "Nuevo" Then
                     EliminaDoc (cNumInve)
                     Sproceso = ""
                     Gridview.Visible = True
                     Call DesHabilita
                  Else
                     Sproceso = ""
                     Gridview.Visible = True
                     Call DesHabilita
                  End If
                Else
                     Exit Sub
                End If
           End If
           
        Else
           If Sproceso = "" Then Unload Me
           If Sproceso = "Guardado" Then DesHabilita
        End If
        
End Sub

Private Sub Form_Load()
Dim RSQL As String
central Me
Carga_Almacen
CargaVista
Sproceso = ""
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub Carga_Almacen()
Dim RSQL As String
Dim I As Integer
RSQL = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open RSQL, VGCNx, adOpenStatic
Do While Not rs.EOF
     Combo1.AddItem (rs(1))
     rs.MoveNext
     If rs.EOF Then Exit Do
Loop
rs.MoveFirst
For I = 0 To rs.RecordCount - 1
  If rs(0) = VGAlma Then
    Combo1.ListIndex = I
    Exit For
  Else
    rs.MoveNext
  End If
Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      txing = ""
      txing.Visible = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      txing = ""
      txing.Visible = False
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      txing = ""
      txing.Visible = False
End Sub

Private Sub Frame4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
      txing = ""
      txing.Visible = False
End Sub

Private Sub Grid_Click()
  Grid.Col = 4
  Call ViewBox
End Sub

Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If OptArti.Value Then
         VGForm1 = 16
         FormAyuArt1.Show 1
         If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
              MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
              Exit Sub
         End If
         If Text1 <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
         End If
Else
     If OptFami.Value Then
        Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
        frmReferencia.Label1.Caption = "Familias de Artículos"
        frmReferencia.Show vbModal
        Adodc2.Close
        If vGUtil(1) <> "" Then
           Text1 = (vGUtil(1))
        End If
        
        If Text1 <> "" Then
           Text2.Enabled = True
           Text2.SetFocus
        End If

     End If
End If

If Text2 <> "" Then Text4_KeyPress (13)

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4_KeyPress (13)
End Sub

Private Sub Text2_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset

If OptArti.Value Then
   VGForm1 = 16
   FormAyuArt1.Show 1
Else
    If OptFami.Value Then
       Adodc2.Open "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA", VGCNx, adOpenStatic, adLockOptimistic
       frmReferencia.Conectar Adodc2, "SELECT FAM_CODIGO,FAM_NOMBRE FROM FAMILIA"
       frmReferencia.Label1.Caption = "Familias de Artículos"
       frmReferencia.Show vbModal
       Adodc2.Close
       If vGUtil(1) <> "" Then
          Text2 = (vGUtil(1))
       End If
       
        If Text1 <> "" Then
           Text2.Enabled = True
           Text2.SetFocus
        End If
       
    End If
End If
    
If Text2 <> "" Then Text4_KeyPress (13)
End Sub

Function NuevoInventa() As Long
    cNumInve = AutoGencodigo()
    If LoadGridNuevo("") = -1 Then
       MsgBox "No Existe referencias  de Articulos con movimientos en Este Almacen  ", vbInformation, "No Hay Información"
       NuevoInventa = -1
       Exit Function
    Else
        Call SaveInvFisico
        Call LoadGrid("", cNumInve)
    End If
    NuevoInventa = 1
End Function

Function AutoGencodigo() As String
    Set rSCod = New ADODB.Recordset
    rSCod.Open "Select max(right(auxnuminve,4)) as May from al_invenfisicocab where auxalma='" & almacen & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
    AutoGencodigo = Format(IIf(IsNull(rSCod!may), 1, rSCod!may + 1), almacen + "000000")
    rSCod.Close
End Function

Sub ViewBox()
  txing.Width = Grid.CellWidth
  txing.Height = Grid.CellHeight - 55
  txing.Top = Grid.Top + Grid.CellTop - 20
  txing.Left = Grid.Left + Grid.CellLeft - 15
  txing.Visible = True
  
  txing.text = Grid.TextMatrix(Grid.Row, 4)
  txing.SetFocus
  txing.SelStart = 0
  txing.SelLength = Len(txing)
  
End Sub
'
'  If Not ExisteElem(0, Vgcnx, "al_invenfisicocab") Then
'        SQL = " Create Table al_invenfisicocab ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXFECH DATETIME ,AUXRESPON TEXT(15),AUXOBSER (40)" & _
'        ", CONSTRAINT Clave PRIMARY KEY ( AUXNUMINVE )  )"
'        Vgcnx.Execute SQL
'  End If
'
'  If Not ExisteElem(0, Vgcnx, "al_invenfisicodet") Then
'        SQL = " Create Table al_invenfisicodet ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXCODART Text(20) ,AUXSTOCK DOUBLE,AUXINGR DOUBLE,AUXDIFE DOUBLE " & _
'        ", CONSTRAINT Clave PRIMARY KEY ( AUXNUMINVE )  )"
'        Vgcnx.Execute SQL
'  End If
'

Function LoadGridNuevo(ByVal busca As String) As Long
    Set rSCod = New ADODB.Recordset
    Dim SQL, criterio, arOrden As String
    
    If CBORDEN.ListIndex <> 0 Then
       arOrden = " MAEART.ACODIGO"
    Else
       arOrden = " MAEART.ADESCRI"
    End If
    
    If busca <> "" Then
       If CBORDEN.ListIndex <> 0 Then
          criterio = " AND MAEART.ACODIGO LIKE '" & busca & "%' "
       Else
          criterio = " AND MAEART.ADESCRI LIKE '" & busca & "%' "
       End If
    Else
      criterio = ""
    End If
    
    If Text1 <> "" And Text2 <> "" Then
       If OptArti.Value = True Then
          criterio = " AND  MAEART.ACODIGO >= '" & Text1 & "' AND MAEART.ACODIGO<='" & Text2 & "'"
       Else
          criterio = " AND  MAEART.AFAMILIA >= '" & Text1 & "' AND MAEART.AFAMILIA<='" & Text2 & "'"
       End If
    End If
    
    SQL = " SELECT MAEART.ACODIGO, MAEART.ADESCRI,MAEART.AFAMILIA, STKART.STSKDIS,0 as Conteo " & _
          " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO WHERE STALMA='" & almacen & "' " & criterio & " order by  " & arOrden

    rSCod.Open SQL, VGCNx, adOpenForwardOnly, adLockReadOnly
    Grid.FixedCols = 0
    If Not rSCod.EOF Then
       Set Grid.DataSource = rSCod
       Grid.FormatString = "^ Código                     |<   Descripción del Articulo                                                 |^ FAMILIA |> Stock Actual |> Ingr. del Conteo"
       LoadGridNuevo = 1
    Else
       sCad = ""
       LoadGridNuevo = -1
    End If
    rSCod.Close
End Function

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   sCad = IIf(KeyAscii <> 8 And KeyAscii <> 13, Chr(KeyAscii), "")
   Text4 = Text4 + sCad
   Text4.SelStart = Len(Text4)
   KeyAscii = IIf(KeyAscii = 8, 8, 0)
   Call LoadGrid(Text4, cNumInve)
End Sub

Private Sub txing_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = 40 Then
        If Grid.Row < Grid.Rows - 1 Then
           Grid.Row = Grid.Row + 1
           Call ViewBox
        Else
           Grid.Row = 1
           Grid.TopRow = 1
           Call ViewBox
        End If
     End If
     If KeyCode = 38 Then
        If Grid.Row > 1 Then
           Grid.Row = Grid.Row - 1
           Call ViewBox
        Else
           Grid.Row = 1
           Call ViewBox
        End If
     End If
     
End Sub

Private Sub txing_KeyPress(KeyAscii As Integer)
     If KeyAscii = 27 Then
        txing = "": txing.Visible = False
     End If

     If KeyAscii = 13 Then
        
        If Not IsNumeric(txing) Then
           MsgBox "Ingrese Valores Numericos...", vbInformation, "Error en el Ingreso"
           Call ViewBox
           Exit Sub
        End If
        Grid.TextMatrix(Grid.Row, 4) = Format(txing, "###,##0.00")
        txing = "": txing.Visible = False
        '*************************
        flagModif = True
        '*************************
        If Grid.Row < Grid.Rows - 1 Then
           Grid.Row = Grid.Row + 1
           Call ViewBox
        Else
           Grid.Row = 1
           Grid.TopRow = 1
           Call ViewBox
        End If
     End If
End Sub
Sub SaveInvFisico()
Dim cCodArt, cFamili As String
Dim CantStk As Double
Dim CantIng As Double
Dim SQL As String
Dim BExist As Boolean
Dim rsExiste As New ADODB.Recordset
rsExiste.Open "Select AUXNUMINVE from al_invenfisicocab where  AUXNUMINVE='" & cNumInve & "'", VGCNx, adOpenForwardOnly, adLockReadOnly
BExist = IIf(Not rsExiste.EOF, True, False)
rsExiste.Close

'RMM****Save Cabezera*****************************************
'Genera o Actualiza Documento
If BExist = True Then
   SQL = "UPDATE al_invenfisicocab SET AUXFECH='" & DTPicker1.Value & "' ,AUXRESPON='" & VGUsuario & "',AUXOBSER='" & cObs & "' WHERE AUXNUMINVE='" & cNumInve & "'"
Else
   SQL = "Insert into al_invenfisicocab (AUXNUMINVE,AUXALMA,AUXFECH,AUXRESPON,AUXOBSER) Values ('" & cNumInve & "','" & almacen & "','" & Format(DTPicker1.Value, "dd/mm/yyyy") & "','" & VGUsuario & "','" & cObs + " " & "')"
End If
VGCNx.Execute SQL
'*************************************************************
'RMM****Save Detalle*****************************************
PBar.Max = Grid.Rows + 50
PBar.Min = 0
Frame5.Visible = True
    For n = 1 To Grid.Rows - 1
        Grid.Row = n
        '*********************
        PBar.Value = n
        Frame5.Refresh
        If Sproceso = "Nuevo" Then
           lbmsg.Caption = "Generando Documento......."
        Else
           lbmsg.Caption = "Guardando el Documento......."
        End If
        '*********************
        cCodArt = Grid.TextMatrix(Grid.Row, 0)
        cFamili = Grid.TextMatrix(Grid.Row, 2)
        CantStk = Val(Grid.TextMatrix(Grid.Row, 3))
        CantIng = Val(Grid.TextMatrix(Grid.Row, 4))
        
        If BExist = True Then
           SQL = "Update al_invenfisicodet Set AUXINGR=" & CantIng & ",AUXDIFE=0 Where AUXNUMINVE='" & cNumInve & "' and AUXCODART='" & cCodArt & "'"
        Else
           SQL = "Insert into al_invenfisicodet (AUXNUMINVE,AUXALMA,AUXCODART,AUXSTOCK,AUXINGR,AUXDIFE,AUXFAMIL) Values ('" & cNumInve & "','" & almacen & "','" & cCodArt & "'," & CantStk & "," & CantIng & ",0,'" & cFamili & "')"
        End If
        VGCNx.Execute SQL
    Next
    
Frame5.Visible = False

End Sub

Function LoadGrid(ByVal busca As String, ByVal xdoc As String) As Long
    Set rSCod = New ADODB.Recordset
    Dim SQL, criterio, arOrden As String
    
    If CBORDEN.ListIndex <> 0 Then
       arOrden = " MAEART.ACODIGO"
    Else
       arOrden = " MAEART.ADESCRI"
    End If
    
    If busca <> "" Then
       If CBORDEN.ListIndex = 1 Then
          criterio = " AND MAEART.ACODIGO LIKE '" & busca & "%' "
       Else
          criterio = " AND MAEART.ADESCRI LIKE '" & busca & "%' "
       End If
    Else
      criterio = ""
    End If
    
    If Text1 <> "" And Text2 <> "" Then
       If OptArti.Value = True Then
          criterio = criterio & " AND  MAEART.ACODIGO >= '" & Text1 & "' AND MAEART.ACODIGO<='" & Text2 & "'"
       Else
          criterio = criterio & " AND  MAEART.AFAMILIA >= '" & Text1 & "' AND MAEART.AFAMILIA<='" & Text2 & "'"
       End If
    End If
    
    'SQL = " SELECT MAEART.ACODIGO, MAEART.ADESCRI,MAEART.AFAMILIA, FORMAT(STKART.STSKDIS,'###,##0.00') ,0 as Conteo " & _
          " FROM MAEART INNER JOIN STKART ON MAEART.ACODIGO = STKART.STCODIGO WHERE STALMA='" & almacen & "' " & Criterio & " order by  " & arOrden
    
    If flagModif = True Then
       MsgBox "Guarde los Datos que usted Modifico para Cambiar el Criterio de Ingreso", vbInformation, "Guardar Datos...!"
       Exit Function
    End If
    
    SQL = " SELECT MAEART.ACODIGO, MAEART.ADESCRI, MAEART.AFAMILIA, al_invenfisicodet.AUXSTOCK,al_invenfisicodet.AUXINGR  " & _
          " FROM MAEART INNER JOIN al_invenfisicodet ON MAEART.ACODIGO = al_invenfisicodet.AUXCODART Where AUXNUMINVE='" & xdoc & "' AND AUXALMA='" & almacen & "' " & criterio & " order by  " & arOrden

    rSCod.Open SQL, VGCNx, adOpenForwardOnly, adLockReadOnly
    Grid.FixedCols = 0
    Grid.Rows = 2
    If Not rSCod.EOF Then
       Set Grid.DataSource = rSCod
       Grid.FormatString = "< Código                     |<   Descripción del Articulo                                                 |^ FAMILIA |>  Stock Actual |> Ingr. del Conteo "
       LoadGrid = 1
    Else
       sCad = ""
       LoadGrid = -1
       Grid.FormatString = "< Código                     |<   Descripción del Articulo                                                 |^ FAMILIA |>  Stock Actual |> Ingr. del Conteo "
       Grid.TextMatrix(1, 0) = "": Grid.TextMatrix(1, 1) = "": Grid.TextMatrix(1, 2) = "": Grid.TextMatrix(1, 3) = "": Grid.TextMatrix(1, 4) = ""
    End If
    rSCod.Close
End Function

Sub EliminaDoc(ByVal arNumdoc As String)
    VGCNx.Execute "Delete from al_invenfisicodet where AUXNUMINVE='" & arNumdoc & "' and AUXALMA='" & almacen & "'"
    VGCNx.Execute "Delete from al_invenfisicocab where AUXNUMINVE='" & arNumdoc & "' and AUXALMA='" & almacen & "'"
End Sub
'  If Not ExisteElem(0, Vgcnx, "al_invenfisicodet") Then
'        SQL = " Create Table al_invenfisicodet ( AUXNUMINVE TEXT(10), AUXALMA Text(3),AUXCODART Text(20) ,AUXSTOCK DOUBLE,AUXINGR DOUBLE,AUXDIFE DOUBLE " & _
'        ", CONSTRAINT Clave PRIMARY KEY ( AUXNUMINVE )  )"
'        Vgcnx.Execute SQL
'  End If

Sub CargaVista()
Dim SQL As String
Set rsInve = New ADODB.Recordset
    SQL = " SELECT al_invenfisicocab.AUXNUMINVE, al_invenfisicocab.AUXFECH, al_invenfisicocab.AUXALMA FROM al_invenfisicocab  where AUXALMA='" & almacen & "'"

    rsInve.Open SQL, VGCNx, adOpenDynamic, adLockReadOnly
    Gridview.FixedCols = 0
    Gridview.Cols = 3
    Gridview.Rows = 2
    If Not rsInve.EOF Then
       Set Gridview.DataSource = rsInve
       Gridview.FormatString = "^ Nro Inventario           |<   Fecha              |^ Almacen     "
    Else
       Gridview.FormatString = "^ Nro Inventario           |<   Fecha              |^ Almacen     "
       Gridview.TextMatrix(1, 0) = "": Gridview.TextMatrix(1, 1) = "": Gridview.TextMatrix(1, 2) = "" ': Gridview.TextMatrix(1, 3) = ""
    End If
    
End Sub

Sub LoadCabezera(ByVal XNum As String)
Dim SQL As String
Set Rs2 = New ADODB.Recordset
Rs2.Open "Select * from al_invenfisicocab where AUXNUMINVE='" & XNum & "'", VGCNx, adOpenForwardOnly
If Not Rs2.EOF Then
   cNumInve = Rs2!AUXNUMINVE
   DTPicker1.Value = Rs2!AUXFECH
   cObs = cNull(Rs2!AUXOBSER)
End If
Rs2.Close
End Sub
Sub Habilita()
     cmdNew.Enabled = False
     cmdModif.Enabled = False
     CMDBORRA.Enabled = False
     Cmdgrabar.Enabled = True
     Frame1.Enabled = True
     Frame2.Enabled = True
     Frame3.Visible = False
     Combo1.Enabled = False
     cmdPrint.Enabled = True
End Sub

Sub DesHabilita()
     cmdNew.Enabled = True
     cmdModif.Enabled = True
     CMDBORRA.Enabled = True
     Cmdgrabar.Enabled = False
     Frame1.Enabled = False
     Frame2.Enabled = False
     Frame3.Visible = True
     cmdPrint.Enabled = False
     cNumInve = ""
     DTPicker1.Value = Format(Now, "dd/mm/yyyy")
     Text1 = ""
     Text2 = ""
     cObs = ""
     Text4 = ""
     Combo1.Enabled = True
End Sub

