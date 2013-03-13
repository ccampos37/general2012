VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInformeInventarioFisico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes de Inventario Fisico"
   ClientHeight    =   6915
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9960
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   9960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   5736
      Left            =   72
      TabIndex        =   23
      Top             =   972
      Width           =   9810
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Gridview 
         Height          =   5085
         Left            =   4035
         TabIndex        =   24
         Top             =   360
         Width           =   5610
         _ExtentX        =   9895
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
         Picture         =   "frmInformeInventarioFisico.frx":0000
         Stretch         =   -1  'True
         Top             =   360
         Width           =   3180
      End
   End
   Begin VB.ComboBox CBORDEN 
      Height          =   288
      ItemData        =   "frmInformeInventarioFisico.frx":2F41
      Left            =   7056
      List            =   "frmInformeInventarioFisico.frx":2F4B
      TabIndex        =   8
      Text            =   "Por Descripción"
      Top             =   2736
      Width           =   2100
   End
   Begin VB.TextBox Text4 
      Height          =   288
      Left            =   1908
      TabIndex        =   7
      Top             =   2736
      Width           =   3396
   End
   Begin VB.Frame Frame4 
      Height          =   912
      Left            =   108
      TabIndex        =   17
      Top             =   0
      Width           =   9555
      Begin VB.CommandButton cmdModif 
         Caption         =   "&Consultar"
         Height          =   636
         Left            =   180
         Picture         =   "frmInformeInventarioFisico.frx":2F6C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   180
         Width           =   775
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   636
         Left            =   1044
         Picture         =   "frmInformeInventarioFisico.frx":33AE
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   180
         Width           =   775
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInformeInventarioFisico.frx":37F0
         Left            =   6330
         List            =   "frmInformeInventarioFisico.frx":37F2
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   324
         Width           =   2976
      End
      Begin VB.CommandButton Command7 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   636
         Left            =   1908
         Picture         =   "frmInformeInventarioFisico.frx":37F4
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "Almacen"
         ForeColor       =   &H80000008&
         Height          =   372
         Left            =   5220
         TabIndex        =   19
         Top             =   360
         Width           =   732
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Criterio de Ingreso"
      Enabled         =   0   'False
      Height          =   1704
      Left            =   5640
      TabIndex        =   10
      Top             =   972
      Width           =   3990
      Begin VB.CheckBox Check1 
         Caption         =   "Mostrar Solo Diferencias"
         Height          =   192
         Left            =   288
         TabIndex        =   30
         Top             =   1404
         Width           =   3252
      End
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
      Height          =   1704
      Left            =   72
      TabIndex        =   9
      Top             =   972
      Width           =   5580
      Begin VB.TextBox cObs 
         Height          =   648
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
         Format          =   49545217
         CurrentDate     =   37090
      End
      Begin VB.Label Label8 
         Caption         =   "Observaciones :"
         Height          =   192
         Left            =   144
         TabIndex        =   22
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
      Height          =   3570
      Left            =   75
      TabIndex        =   0
      Top             =   3135
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   6297
      _Version        =   393216
      FixedCols       =   0
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
   Begin VB.Frame Frame5 
      Height          =   1164
      Left            =   2592
      TabIndex        =   25
      Top             =   3132
      Visible         =   0   'False
      Width           =   4512
      Begin MSComctlLib.ProgressBar PBar 
         Height          =   228
         Left            =   324
         TabIndex        =   26
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
         TabIndex        =   27
         Top             =   360
         Width           =   3900
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Orden :"
      Height          =   192
      Left            =   5868
      TabIndex        =   21
      Top             =   2808
      Width           =   984
   End
   Begin VB.Label Label7 
      Caption         =   "Buscar  :"
      Height          =   228
      Left            =   108
      TabIndex        =   20
      Top             =   2772
      Width           =   1704
   End
End
Attribute VB_Name = "frmInformeInventarioFisico"
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

Private Sub cmdModif_Click()
     Frame5.Visible = True
     Sproceso = "Ver"
     Frame5.Refresh
     If Gridview.Rows = 2 And Gridview.TextMatrix(Gridview.Row, 0) = "" Then Exit Sub
     Call LoadCabezera(Gridview.TextMatrix(Gridview.Row, 0))
     Frame5.Refresh
     Call LoadGrid("", Gridview.TextMatrix(Gridview.Row, 0))
     Call Habilita
     Frame5.Visible = False
End Sub



Private Sub cmdPrint_Click()
Dim CTIME As String
Dim ccadena As String
CTIME = Format(Time, "hh:mm:ss")

If OptArti.Value = True Then
   CrystalReport1.WindowTitle = "Inv509 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv509.Rpt"
Else
   CrystalReport1.WindowTitle = "Inv508 -- Control de Inventarios"
   CrystalReport1.ReportFileName = VGParamSistem.RutaReport & "inv508.Rpt"
End If

If OptArti.Value = True Then
   If Text1 = "" Or Text2 = "" Then
      ccadena = " {al_invenfisicodet.AUXNUMINVE}='" & cNumInve & "' AND {al_invenfisicodet.AUXALMA}='" & almacen & "'"
   Else
      ccadena = " {al_invenfisicodet.AUXNUMINVE}='" & cNumInve & "' AND {al_invenfisicodet.AUXALMA}='" & almacen & "' AND {al_invenfisicodet.AUXALMA}='" & almacen & "' AND {al_invenfisicodet.AUXCODART}>='" & Text1 & "' and {al_invenfisicodet.AUXCODART}<='" & Text2 & "'"
   End If
Else
   If Text1 = "" Or Text2 = "" Then
      ccadena = " {al_invenfisicodet.AUXNUMINVE}='" & cNumInve & "' AND {al_invenfisicodet.AUXALMA}='" & almacen & "'"
   Else
      ccadena = " {al_invenfisicodet.AUXNUMINVE}='" & cNumInve & "' AND {al_invenfisicodet.AUXALMA}='" & almacen & "' AND {al_invenfisicodet.AUXFAMIL}>='" & Text1 & "' and {al_invenfisicodet.AUXFAMIL}<='" & Text2 & "'"
   End If
End If

If Check1.Value = 1 Then
   ccadena = ccadena & " AND {@XDIFEREN}<>0"
End If

CrystalReport1.WindowState = crptMaximized
Call Ubi_Tab(CrystalReport1)
CrystalReport1.WindowShowPrintBtn = True
CrystalReport1.WindowShowRefreshBtn = True
CrystalReport1.WindowShowSearchBtn = True
CrystalReport1.WindowShowPrintSetupBtn = True
CrystalReport1.formulas(0) = "Hora = '" & CTIME & "'"
CrystalReport1.formulas(1) = "Empresa = '" & Mid(VGparametros.RucEmpresa, 1, 20) & "'"
CrystalReport1.formulas(2) = "XFECHINVE='" & Format(Now, "DD/MM/YYYY") & "'"
CrystalReport1.formulas(3) = ""
'CrystalReport1.SelectionFormula = ccadena
CrystalReport1.ReplaceSelectionFormula (ccadena)
CrystalReport1.WindowTop = 100
CrystalReport1.WindowLeft = 150
CrystalReport1.DiscardSavedData = True
CrystalReport1.Destination = crptToWindow
 If CrystalReport1.Status <> 2 Then CrystalReport1.Action = 1
End Sub

Private Sub Combo1_Click()
rs.MoveFirst
rs.Move Combo1.ListIndex
almacen = Format(rs(0), "00")
CargaVista
End Sub



Private Sub Command7_Click()
   If Sproceso = "Ver" Then
      DesHabilita
      Sproceso = ""
   Else
      Unload Me
   End If
End Sub

Private Sub Form_Load()
Dim rsql As String
central Me
Carga_Almacen
CargaVista
Sproceso = ""
DTPicker1.Value = Format(Now, "DD/MM/YYYY")
End Sub

Private Sub Carga_Almacen()
Dim rsql As String
Dim I As Integer
rsql = "Select TAALMA,TADESCRI FROM TabAlm "
Set rs = New ADODB.Recordset
rs.Open rsql, VGCNx, adOpenStatic
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



Private Sub Text1_DblClick()
Dim Adodc2 As ADODB.Recordset
Set Adodc2 = New ADODB.Recordset
If OptArti.Value Then
         VGForm1 = 20
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
   VGForm1 = 20
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




Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4_KeyPress (13)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   sCad = IIf(KeyAscii <> 8 And KeyAscii <> 13, Chr(KeyAscii), "")
   Text4 = Text4 + sCad
   Text4.SelStart = Len(Text4)
   If KeyAscii = 13 Then
      KeyAscii = IIf(KeyAscii = 8, 8, 0)
      Call LoadGrid(Text4, cNumInve)
   Else
      KeyAscii = IIf(KeyAscii = 8, 8, 0)
   End If
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
    
    SQL = " SELECT MAEART.ACODIGO, MAEART.ADESCRI, MAEART.AFAMILIA, al_invenfisicodet.AUXSTOCK,al_invenfisicodet.AUXINGR " & _
          " FROM MAEART INNER JOIN al_invenfisicodet ON MAEART.ACODIGO = al_invenfisicodet.AUXCODART Where AUXNUMINVE='" & xdoc & "' AND AUXALMA='" & almacen & "' " & criterio & " order by  " & arOrden

    rSCod.Open SQL, VGCNx, adOpenForwardOnly, adLockReadOnly
    grid.FixedCols = 0
    grid.Rows = 2
    If Not rSCod.EOF Then
       Set grid.DataSource = rSCod
       grid.FormatString = "< Código                     |<   Descripción del Articulo                                                 |^ FAMILIA |>   Stock Actual  |>  Ingr. del Conteo  "
       LoadGrid = 1
    Else
       sCad = ""
       LoadGrid = -1
       grid.FormatString = "< Código                     |<   Descripción del Articulo                                                 |^ FAMILIA |>   Stock Actual  |>  Ingr. del Conteo  "
       grid.TextMatrix(1, 0) = "": grid.TextMatrix(1, 1) = "": grid.TextMatrix(1, 2) = "": grid.TextMatrix(1, 3) = "": grid.TextMatrix(1, 4) = ""
    End If
    rSCod.Close
End Function


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
       Gridview.FormatString = "^ Nro Inventario             |<   Fecha                    |^ Almacen         "
    Else
       Gridview.FormatString = "^ Nro Inventario             |<   Fecha                    |^ Almacen         "
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
     'cmdNew.Enabled = False
     cmdModif.Enabled = False
     'cmdBorra.Enabled = False
     'CmdGrabar.Enabled = True
     'Frame1.Enabled = True
     Frame2.Enabled = True
     Frame3.Visible = False
     Combo1.Enabled = False
     cmdPrint.Enabled = True
End Sub

Sub DesHabilita()
     'cmdNew.Enabled = True
     cmdModif.Enabled = True
     'cmdBorra.Enabled = True
     'CmdGrabar.Enabled = False
     Frame1.Enabled = False
     Frame2.Enabled = False
     Frame3.Visible = True
     cmdPrint.Enabled = False
     DTPicker1.Value = Format(Now, "dd/mm/yyyy")
     Text1 = ""
     Text2 = ""
     cObs = ""
     Text4 = ""
     Combo1.Enabled = True
End Sub

