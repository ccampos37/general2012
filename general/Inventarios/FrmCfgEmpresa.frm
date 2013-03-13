VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmCfgEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de Empresas "
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "FrmCfgEmpresa.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7455
   Begin VB.Frame Frame3 
      Height          =   1080
      Left            =   60
      TabIndex        =   20
      Top             =   3345
      Width           =   7020
      Begin VB.CommandButton CmdEli2 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2430
         Picture         =   "FrmCfgEmpresa.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdAgre 
         Caption         =   "&Agregar"
         Height          =   675
         Left            =   3525
         Picture         =   "FrmCfgEmpresa.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   6000
         Picture         =   "FrmCfgEmpresa.frx":1016
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1350
         Picture         =   "FrmCfgEmpresa.frx":1458
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   270
         Picture         =   "FrmCfgEmpresa.frx":189A
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdCon 
         Caption         =   "&Consulta"
         Height          =   675
         Left            =   4635
         Picture         =   "FrmCfgEmpresa.frx":1CDC
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   225
         Width           =   775
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2400
         Picture         =   "FrmCfgEmpresa.frx":211E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   5280
         Picture         =   "FrmCfgEmpresa.frx":2560
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   225
         Visible         =   0   'False
         Width           =   775
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmCfgEmpresa.frx":29A2
      Height          =   3135
      Left            =   105
      TabIndex        =   16
      Top             =   135
      Width           =   7110
      _ExtentX        =   12541
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   2
      WrapCellPointer =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "EMP_CODIGO"
         Caption         =   "   Empresa"
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
         DataField       =   "EMP_RAZON_NOMBRE"
         Caption         =   "                                      Razón Social"
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
         DataField       =   "EMP_RUC_DOCUMENTO"
         Caption         =   "        R.U.C."
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
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4185.071
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1365.165
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3120
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   7110
      Begin VB.TextBox TxDire 
         Height          =   285
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "TxDire"
         Top             =   975
         Width           =   4695
      End
      Begin VB.TextBox TxReporte 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1950
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2625
         Width           =   4635
      End
      Begin VB.TextBox TxPantalla 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1950
         MaxLength       =   20
         TabIndex        =   7
         Top             =   2295
         Width           =   4665
      End
      Begin VB.TextBox TxRep 
         Height          =   285
         Left            =   1950
         MaxLength       =   40
         TabIndex        =   6
         Text            =   "Txcon"
         Top             =   1965
         Width           =   4665
      End
      Begin VB.TextBox Txfax 
         Height          =   285
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Txfax"
         Top             =   1635
         Width           =   2175
      End
      Begin VB.TextBox Txtel 
         Height          =   285
         Left            =   1950
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "Txtel"
         Top             =   1305
         Width           =   2175
      End
      Begin VB.TextBox TxRazon 
         Height          =   285
         Left            =   1950
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "TxRazon"
         Top             =   645
         Width           =   4695
      End
      Begin VB.TextBox TxRuc 
         Height          =   285
         Left            =   5055
         MaxLength       =   11
         TabIndex        =   1
         Text            =   "TxRuc"
         Top             =   300
         Width           =   1575
      End
      Begin VB.TextBox TxCod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1950
         MaxLength       =   3
         TabIndex        =   0
         Text            =   "Tx"
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Dirección                :"
         Height          =   255
         Left            =   300
         TabIndex        =   19
         Top             =   990
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Reporte                  :"
         Height          =   255
         Index           =   13
         Left            =   315
         TabIndex        =   18
         Top             =   2655
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Pantalla                  :"
         Height          =   255
         Index           =   12
         Left            =   300
         TabIndex        =   17
         Top             =   2325
         Width           =   1635
      End
      Begin VB.Label Label12 
         Caption         =   "Representante       :"
         Height          =   255
         Left            =   300
         TabIndex        =   15
         Top             =   1995
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "Fax                         :"
         Height          =   255
         Left            =   300
         TabIndex        =   14
         Top             =   1650
         Width           =   1770
      End
      Begin VB.Label Label9 
         Caption         =   "Teléfono                 :"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   1335
         Width           =   1560
      End
      Begin VB.Label Label4 
         Caption         =   "R.U.C.  "
         Height          =   255
         Left            =   4260
         TabIndex        =   12
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Razón Social          :"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   675
         Width           =   1620
      End
      Begin VB.Label Label1 
         Caption         =   "Código                    :    "
         Height          =   255
         Index           =   0
         Left            =   315
         TabIndex        =   10
         Top             =   345
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmCfgEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim nTipOper As Integer
Dim nTra As Integer
Dim csql As String
Dim CSQL2 As String
Dim cSql3 As String
Dim cSql4 As String
Dim cCod As String
Dim nTra2 As Integer
Dim cBase As String

Private Sub CmdAgre_Click()
    If adodc1.EOF Or adodc1.BOF Then Exit Sub
    If Not AGREGARBASE(adodc1(0)) Then Screen.MousePointer = 1: Exit Sub
'    Set cConexCom = New ADODB.Connection  'BD. Común
'    cConexCom.CursorLocation = adUseClient
'    cConexCom.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRuta2 & ";"
'    cConexCom.Open
    DataGrid1.SetFocus
    DataGrid1_RowColChange 0, 0
End Sub

Private Sub CmdCon_Click()
If adodc1.RecordCount > 0 Then
    cCod = adodc1("EMP_CODIGO")
    Frame1.Caption = "Consulta de Empresa"
    OculObj (False)
    Limpiar1
    Mostrar
    InhObj (False)
    CmdGrabar.Enabled = False
    CmdGrabar.Visible = True
    CmdSalir2.Visible = True
    Frame1.Visible = True
End If
End Sub

Private Sub CmdEli2_Click()
Dim nNd As Integer
On Error GoTo ElErr
'Esta opcion puede estar inhabilitada si contabilidad esta instalada
If adodc1.RecordCount > 0 Then
    cCod = adodc1("EMP_CODIGO")
    If VGCOMP = cCod Then
        MsgBox "No se puede Eliminar la Empresa Activa", vbInformation, "Inventarios"
        Exit Sub
    End If
    If MsgBox("Advertencia ! Antes de eliminar la empresa verifique que no exista informacion " & Chr(13) & _
               "de otros sistemas de Enterprise Solutions " & Chr(13) & _
               "si procede a eliminar perdera la información de estos sistemas ", vbExclamation + vbOKCancel) = vbCancel Then
       Exit Sub
    End If
    If MsgBox("Desea Eliminar la Empresa", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        nNd = Pos_Dato(adodc1)
        csql = "Delete From EmpresA where EMP_CODIGO = '" & cCod & "'"
        CSQL2 = "Delete From Parametros where CTCCOD = '" & cCod & "'"
'       cSql3 = "Delete From Usuario_Fac where CTEMPRESA = '" & cCod & "'"
        cSql4 = "Delete From Punto_Venta  where PV_EMPRESA = '" & cCod & "'"
        nTra = 1
        cn.BeginTrans
        
        If nNd <> 0 Then adodc1.AbsolutePosition = nNd
            Call Borrar_Empresa(cCod)
            DataGrid1.SetFocus
        End If
        cn.Execute CSQL2
 '       cn.Execute cSql3
'        cn.Execute cSql4
        cn.CommitTrans
        nTra = 0
        adodc1.Requery
        DataGrid1_RowColChange 0, 0
    End If
Exit Sub
ElErr:
    MsgBox Err.Description
    If nTra = 1 Then cn.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()
On Error GoTo GraErr
'Esta opcion puede estar inhabilitada si contabilidad esta instalada
If nTipOper = 1 Then
    TxCod = Format(TxCod, "000")
   If Existe(2, TxCod, "Empresa", "EMP_CODIGO", False) Then
        MsgBox "La Empresa ya existe", vbInformation, "Mensaje"
        TxCod.SetFocus
        Exit Sub
    End If
End If

If nTipOper = 1 Or nTipOper = 2 Then
    If Trim(TxRazon) = "" Then
        MsgBox "Ingrese Razón Social", vbInformation, "Mensaje"
        TxRazon.SetFocus: Exit Sub
    End If
    If Trim(TxRuc) = "" Then
        MsgBox "Ingrese R.U.C. de la Empresa", vbInformation, "Inventarios"
        TxRuc.SetFocus: Exit Sub
    ElseIf Validar_RUC(TxRuc) = False Then
        TxRuc.SetFocus: Exit Sub
    End If
    If Trim(TxPantalla) = "" Then
        MsgBox "Ingrese el dato de Pantalla", vbInformation, "Inventarios"
        TxPantalla.SetFocus: Exit Sub
    End If
    If Trim(TxReporte) = "" Then
        MsgBox "Ingrese dato del Reporte", vbInformation, "Inventarios"
        TxReporte.SetFocus: Exit Sub
    End If
End If

If nTipOper = 1 Then
    csql = "Insert Into EmpresA (EMP_CODIGO,EMP_RAZON_NOMBRE,EMP_RUC_DOCUMENTO,"
    csql = csql & "EMP_DIRECCION,EMP_TELEFONO,EMP_FAX,EMP_REPRESENTANTE,EMP_REPORTE,EMP_PANTALLA) "
    csql = csql & " Values ('" & TxCod & "','" & TxRazon & "','" & TxRuc & "','" & TxDire & "',"
    csql = csql & "'" & Txtel & "','" & Txfax & "','" & TxRep & "','" & TxReporte & "','" & TxPantalla & "')"
    
    CSQL2 = "Insert Into Parametros (CTCCOD) Values ('" & TxCod & "')"
    
    nTra = 1
    cn.BeginTrans
    cn.Execute csql
    cn.Execute CSQL2
    cn.CommitTrans
    nTra = 0
    

    
ElseIf nTipOper = 2 Then
    csql = "Update EmpresA Set "
    csql = csql & "EMP_RAZON_NOMBRE='" & TxRazon & "',EMP_RUC_DOCUMENTO ='" & TxRuc & "',"
    csql = csql & "EMP_DIRECCION = '" & TxDire & "',EMP_TELEFONO = '" & Txtel & "', EMP_FAX = '" & Txfax & "',"
    csql = csql & "EMP_REPRESENTANTE = '" & TxRep & "',EMP_REPORTE = '" & TxReporte & "',  EMP_PANTALLA = '" & TxPantalla & "'   Where EMP_CODIGO = '" & TxCod & "'"
    
    nTra = 1
    cn.BeginTrans
    cn.Execute csql
    cn.CommitTrans
    nTra = 0
End If
adodc1.Requery


If nTipOper = 1 Then
    Limpiar1
    TxCod.SetFocus
ElseIf nTipOper = 2 Or nTipOper = 3 Then
    CmdSalir2_Click
End If
Exit Sub
GraErr:
    MsgBox Err.Description
    If nTra = 1 Then cn.RollbackTrans
    Exit Sub
Carp:
End Sub

Private Sub CmdIng_Click()
'If UCase(Dir$(Left(cRuta6, Len(cRuta6) - 11) & "DATA\BDENCOT.MDB", vbArchive)) = "BDENCOT.MDB" Then
'    MsgBox "Usted tiene el sistema de contabilidad!. " & Chr(13) & " Por lo tanto, mientras tenga el sistema contable solamente podra crear empresas desde alli!", vbInformation
'    Exit Sub
'End If
Frame1.Caption = "Ingreso de Empresa"
OculObj (False)
Limpiar1
nTipOper = 1
CmdGrabar.Visible = True
CmdGrabar.Enabled = True
CmdSalir2.Visible = True
Frame1.Visible = True
TxCod.SetFocus
End Sub

Private Sub CmdModi_Click()
If adodc1.RecordCount > 0 Then
    cCod = adodc1("EMP_CODIGO")
    Frame1.Caption = "Modificación de Empresas"
    OculObj (False)
    Limpiar1
    Mostrar
    nTipOper = 2
    CmdGrabar.Visible = True
    CmdGrabar.Enabled = True
    CmdSalir2.Visible = True
    Frame1.Visible = True
    TxRazon.SetFocus
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()
CmdGrabar.Visible = False
CmdSalir2.Visible = False
Frame1.Visible = False
OculObj (True)
InhObj (True)
TxCod.Enabled = True
If DataGrid1.Enabled And DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If adodc1.EOF Then Exit Sub
End Sub

Private Sub Form_Activate()
If DataGrid1.Enabled And DataGrid1.Visible Then DataGrid1.SetFocus
End Sub

Private Sub Form_Load()
central Me

Set adodc1 = New ADODB.Recordset
adodc1.Open "Select EMP_CODIGO,EMP_RAZON_NOMBRE,EMP_RUC_DOCUMENTO  From Empresa order by EMP_CODIGO", cn, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh

''Verificar si Contabilidad esta instalado, y se inhabilitara algunos controles (Ingreso,Eliminación,Modificación)
'cBase = cRuta4
'If UCase(Dir$(cBase)) = UCase(VGNameCont & ".MDB") Then
'    CmdIng.Enabled = False
'    CmdModi.Enabled = False
'    CmdEli2.Enabled = False
'    If UCase(Dir$(sName & "Data\" & VGCOMP & "\BDCOMUN.MDB")) = "BDCOMUN.MDB" Then
'        CmdAgre.Enabled = False
'    Else
'        CmdAgre.Enabled = True
'    End If
'Else

'    If UCase(Dir$(sName & "Data\" & VGCOMP & "\BDCOMUN.MDB")) = "BDCOMUN.MDB" Then
'            CmdAgre.Enabled = False
'        Else
'            CmdAgre.Enabled = True
'    End If
    CmdIng.Enabled = True
    CmdModi.Enabled = True
    cmdEli2.Enabled = True
'End If
End Sub

Private Sub Limpiar1()
TxCod = "": TxRuc = "": TxRazon = " "
TxDire = " ": Txfax = " ": TxRep = " "
Txtel = " ": TxPantalla = " ": TxReporte = " "
End Sub

Private Sub OculObj(bT As Boolean)
DataGrid1.Visible = bT
CmdIng.Visible = bT
CmdModi.Visible = bT
cmdEli2.Visible = bT
CmdSalir.Visible = bT
CmdCon.Visible = bT
CmdAgre.Visible = bT
End Sub

Private Sub TxCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    TxCod = Format(TxCod, "000")
    If Existe(2, TxCod, "EMPRESA", "EMP_CODIGO", False) = False Then
        SendKeys "{Tab}"
    Else
        MsgBox "La Empresa  ya existe", vbInformation, "Mensaje"
        TxCod.SetFocus
    End If
End If
End Sub

Private Sub TxDire_GotFocus()
Enfoque TxDire
End Sub

Private Sub TxDire_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}"
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub Txfax_GotFocus()
Enfoque Txfax
End Sub

Private Sub Txfax_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxPantalla_GotFocus()
Enfoque TxPantalla
End Sub

Private Sub TxPantalla_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{tab}"
End Sub

Private Sub TxRazon_GotFocus()
Enfoque TxRazon
End Sub

Private Sub TxRazon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxRazon) = "" Then
        MsgBox "Ingrese Razón Social", vbInformation, "Mensaje"
        TxRazon.SetFocus: Exit Sub
    Else
        SendKeys "{Tab}"
    End If
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End If
End Sub

Private Sub TxRep_GotFocus()
Enfoque TxRep
End Sub

Private Sub TxRep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub TxReporte_GotFocus()
Enfoque TxReporte
End Sub

Private Sub TxReporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGrabar.SetFocus
End Sub

Private Sub TxRuc_GotFocus()
Enfoque TxRuc
End Sub

Private Sub TxRuc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Validar_RUC(Trim(TxRuc)) = False Then
        TxRuc.SetFocus: Exit Sub
    End If
    SendKeys "{Tab}"
End If
End Sub

Private Sub Txtel_GotFocus()
Enfoque Txtel
End Sub

Private Sub Txtel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{Tab}"
End Sub

Private Sub Mostrar()
Dim cS1 As String, cR1 As ADODB.Recordset

cS1 = "Select * from EmpresA where EMP_CODIGO = '" & cCod & "'"
Set cR1 = New ADODB.Recordset
cR1.Open cS1, cn, adOpenStatic
If cR1.RecordCount > 0 Then
    TxCod = cR1("EMP_CODIGO")
    TxRuc = IIf(IsNull(cR1("EMP_RUC_DOCUMENTO")), " ", cR1("EMP_RUC_DOCUMENTO"))
    TxRazon = IIf(IsNull(cR1("EMP_RAZON_NOMBRE")), " ", cR1("EMP_RAZON_NOMBRE"))
    TxDire = IIf(IsNull(cR1("EMP_DIRECCION")), " ", cR1("EMP_DIRECCION"))
    Txtel = IIf(IsNull(cR1("EMP_TELEFONO")), " ", cR1("EMP_TELEFONO"))
    Txfax = IIf(IsNull(cR1("EMP_FAX")), " ", cR1("EMP_FAX"))
    TxRep = IIf(IsNull(cR1("EMP_REPRESENTANTE")), " ", cR1("EMP_REPRESENTANTE"))
    TxPantalla = IIf(IsNull(cR1("EMP_PANTALLA")), " ", cR1("EMP_PANTALLA"))
    TxReporte = IIf(IsNull(cR1("EMP_REPORTE")), " ", cR1("EMP_REPORTE"))
    TxCod.Enabled = False
Else
    MsgBox "La Empresa ha sido Eliminada", vbInformation, "Mensaje"
    cR1.Close: Exit Sub
End If
cR1.Close
End Sub

Private Sub InhObj(bT As Boolean)
TxCod.Enabled = bT
TxRuc.Enabled = bT
TxRazon.Enabled = bT
TxDire.Enabled = bT
Txfax.Enabled = bT
Txtel.Enabled = bT
TxRep.Enabled = bT
TxPantalla.Enabled = bT
TxReporte.Enabled = bT
End Sub

Private Sub Borrar_Empresa(cEmp As String)
Dim name As String
On Error GoTo cArch
    'BORRANDO EL BDCOMUN
    name = sName & "Data\" & Trim(cEmp) & "\" & "BDCOMUN.MDB"
    If Dir(name) <> "" Then
        Kill name
    End If
    
    'BORRANDO EL BDTRANSF
    name = sName & "Data\" & Trim(cEmp) & "\" & "BDTransf.mdb"
    If Dir(name) <> "" Then
        Kill name
    End If
    
    name = sName & "Data\" & Trim(cEmp) & "\"
    Dir name
    Exit Sub
cArch:
 MsgBox "Error : " + error + " Verifique la operación", vbInformation, "Mensaje de Error"
 Exit Sub
End Sub

