VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FormAyuArt1 
   Caption         =   "Articulos"
   ClientHeight    =   5250
   ClientLeft      =   2640
   ClientTop       =   1260
   ClientWidth     =   7965
   ControlBox      =   0   'False
   LinkTopic       =   "Form11"
   ScaleHeight     =   5250
   ScaleWidth      =   7965
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      Height          =   4005
      Left            =   120
      TabIndex        =   2
      Top             =   105
      Width           =   7740
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormAyuArt1.frx":0000
         Left            =   5820
         List            =   "FormAyuArt1.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   345
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   3
         Top             =   360
         Width           =   3045
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3075
         Left            =   105
         TabIndex        =   7
         Top             =   840
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   5424
         _Version        =   393216
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
               LCID            =   2058
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
               LCID            =   2058
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
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "FormAyuArt1.frx":0031
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   4860
         TabIndex        =   6
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Campo"
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   5175
      Picture         =   "FormAyuArt1.frx":08FB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4350
      Width           =   930
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   2010
      Picture         =   "FormAyuArt1.frx":0D3D
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4335
      Width           =   930
   End
End
Attribute VB_Name = "FormAyuArt1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim varform As Form
Dim adodc1 As ADODB.Recordset

Private Sub Combo1_Click()
 If Combo1.ListIndex < 0 Then Exit Sub
 If IsNull(adodc1) Then Exit Sub
 If adodc1.State = 0 Then Exit Sub
  adodc1.Sort = adodc1.Fields(Combo1.ListIndex).Name
 'DataGrid1.ReBind
 
End Sub

Private Sub Command1_Click()
If adodc1.RecordCount > 0 Then
    If Not IsNull(DataGrid1.Bookmark) Then
         adodc1.Bookmark = DataGrid1.Bookmark
    End If
    If varform.Text1 <> "" And VGForm1 <> 10 And VGForm1 <> 11 And VGForm1 <> 12 And VGForm1 <> 14 And VGForm1 <> 21 Then
        varform.Text2 = adodc1("ACODIGO")
    ElseIf VGForm1 <> 10 And VGForm1 <> 21 Then
         varform.Text1 = adodc1("ACODIGO")
    End If
    If VGForm1 = 10 Or VGForm1 = 11 Or VGForm1 = 12 Or VGForm1 = 21 Then
         varform.Text1 = adodc1("ACODIGO")
         varform.Label3 = Mid(adodc1("ADESCRI"), 1, 25)
     End If
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Cod As String
'central Me ' Centra el Formulario
AlinearAyuda Me
Set adodc1 = New ADODB.Recordset

Init_ControlDataGrid DataGrid1
 
Cod = ""
Combo1.ListIndex = 1
Select Case VGForm1
    Case 3
                Set varform = FormStkAlm
    Case 4
                Set varform = FormKardexMov
    Case 5
                Set varform = FrmKardex
    Case 6
                Set varform = FormKardexVal
    Case 7
                Set varform = FormArtRep
    Case 8
                Set varform = cstkardexmovi
    Case 9
                Set varform = frmRotacion
    Case 10
                Set varform = FormCasillero
                SQL = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p ORDER BY p.acodigo"
    Case 11
                Set varform = FormMovArt
    Case 12
                Set varform = frmTransMABEL
                SQL = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p ORDER BY p.acodigo"
    Case 14
                Set varform = frmPrueba
    
    Case 16     'RMM*************************
                Set varform = frmRegistroInventarioFisico
                'RMM*************************
    Case 17     'RMM*************************
                Set varform = frmStockLoteSerie
                'RMM*************************
    Case 19     'RMM*************************
                Set varform = FrmKardexValTXDetallado
                'RMM*************************
    Case 20     'RMM*************************
                Set varform = frmInformeInventarioFisico
                'RMM*************************
    Case 21     'RMM*************************
                Set varform = frmReglotes
                'RMM*************************
    Case 22     'RMM*************************
                Set varform = frmArticuloXCenCos
                'RMM*************************
    Case 23     'RMM*************************
                Set varform = frmKardexLote
    Case 30     ' establecimientos
                Set varform = FrmKarValxEst
    Case 31     ' x empresa
                Set varform = FrmKarValdetxEmpresa
                                
    End Select
If VGForm1 <> "10" And VGForm1 <> "12" Then
    If varform.Text1 <> "" And VGForm1 <> 19 And VGForm1 = 30 And VGForm1 = 31 Then
              Cod = varform.Text1
              SQL = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n where  p.ACODIGO =  n.STCODIGO  and p.ACODIGO >= '" & Cod & "'and n.STALMA = '" & VGAlma & "'    ORDER BY ACODIGO " '   group BY p.ACODIGO " ' n.STSKDIS <> 0
    Else
        If VGForm1 = 21 Or VGForm1 = 19 Or VGForm1 = 30 Or VGForm1 = 31 Then
           SQL = "select  ACODIGO, ADESCRI,AUNIDAD,ACODIGO2 from MaeArt "    '     ORDER BY ACODIGO " '   group BY p.ACODIGO "    'n.STSKDIS <>0 and
        Else
           SQL = "select  p.ACODIGO, p.ADESCRI,p.AUNIDAD,p.ACODIGO2 from MaeArt p, StkArt n where   p.ACODIGO =  n.STCODIGO and n.STALMA = '" & VGAlma & "'     ORDER BY ACODIGO " '   group BY p.ACODIGO "    'n.STSKDIS <>0 and
        End If
    End If
End If
adodc1.Open SQL, VGCNx, adOpenStatic
If adodc1.BOF Or adodc1.EOF Then
   Exit Sub
End If
adodc1.MoveFirst
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
CarObj                                ' Objetos
End Sub

Private Sub CarObj()        ' Carga Objetos
 DataGrid1.Columns(0).Locked = True
 DataGrid1.Columns(0).WrapText = True
 DataGrid1.Columns(0).Caption = "   CODIGO"
 DataGrid1.Columns(1).Caption = "       DESCRIPCION"
 DataGrid1.Columns(2).Caption = "   UNIDAD"
 DataGrid1.Columns(3).Caption = "   COD FAB."
 DataGrid1.Columns(0).Width = 2000
 DataGrid1.Columns(1).Width = 4000
 DataGrid1.Columns(2).Width = 1500
 DataGrid1.Columns(3).Width = 1500
End Sub

Private Sub Text1_Change()
 Dim Ant As Long
 Dim citerio As String
  Ant = 0
 'adodc1.Sort
  If adodc1.RecordCount <> 0 Then
        If Text1 <> "" Then
                If Combo1.ListIndex = 0 Then
                        criterio = adodc1.Fields(0).Name & " LIKE '" & UCase(Trim(Text1)) & "*'"
                ElseIf Combo1.ListIndex = 2 Then
                        criterio = adodc1.Fields(3).Name & " LIKE '" & UCase(Trim(Text1)) & "*'"
                Else
                        criterio = adodc1.Fields(1).Name & " LIKE '" & UCase(Trim(Text1)) & "*'"
                End If
                adodc1.Find criterio
                If adodc1.EOF Then
                      adodc1.MoveFirst
                End If
        End If
  End If
End Sub
