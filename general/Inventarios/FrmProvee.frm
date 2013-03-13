VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form FrmProvee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualiza Datos Generales de Proveedor"
   ClientHeight    =   5490
   ClientLeft      =   735
   ClientTop       =   2910
   ClientWidth     =   7515
   Icon            =   "FrmProvee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7515
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   20
      Top             =   4320
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   675
         Left            =   2160
         Picture         =   "FrmProvee.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton CmdSalir2 
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4320
         Picture         =   "FrmProvee.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   775
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   480
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "FrmProvee.frx":114E
      Height          =   3465
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6112
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   17
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "PRVCCODIGO"
         Caption         =   "              CODIGO"
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
         DataField       =   "PRVCNOMBRE"
         Caption         =   "                                  RAZON SOCIAL"
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
         DataField       =   "PRVCDIRECC"
         Caption         =   "                                   DIRECCION"
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
         DataField       =   "PRVCTELEF1"
         Caption         =   "             TELEFONO"
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
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4680
         EndProperty
         BeginProperty Column02 
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   4665.26
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            Locked          =   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   2099.906
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   4155
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   7335
      Begin VB.CheckBox chkAsignarRUC 
         Caption         =   "Asignar a RUC?"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   2175
         MaxLength       =   15
         TabIndex        =   8
         Text            =   "Text7"
         Top             =   2130
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   12
         Text            =   "Text11"
         Top             =   3600
         Width           =   1770
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2175
         MaxLength       =   11
         TabIndex        =   0
         Text            =   "12345678901"
         Top             =   390
         Width           =   1290
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   5640
         MaxLength       =   11
         TabIndex        =   2
         Text            =   "12345678901"
         Top             =   390
         Width           =   1320
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   4
         Text            =   "Text3"
         Top             =   1080
         Width           =   4755
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   5
         Text            =   "Text4"
         Top             =   1440
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   6
         Text            =   "Text5"
         Top             =   1800
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   50
         TabIndex        =   3
         Text            =   "Text2"
         Top             =   720
         Width           =   4755
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   5715
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "Te"
         Top             =   2190
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   2160
         MaxLength       =   15
         TabIndex        =   9
         Text            =   "Text8"
         Top             =   2520
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   2160
         MaxLength       =   40
         TabIndex        =   10
         Text            =   "Text9"
         Top             =   2880
         Width           =   3015
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   2160
         MaxLength       =   20
         TabIndex        =   11
         Text            =   "Text10"
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Teléfono   "
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   2190
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Código "
         Height          =   195
         Left            =   360
         TabIndex        =   39
         Top             =   480
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "RUC   "
         Height          =   195
         Left            =   5160
         TabIndex        =   38
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Dirección "
         Height          =   195
         Left            =   360
         TabIndex        =   37
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label Label5 
         Caption         =   "Giro del Proveedor "
         Height          =   195
         Left            =   4110
         TabIndex        =   35
         Top             =   2205
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "Nombre "
         Height          =   195
         Left            =   360
         TabIndex        =   34
         Top             =   810
         Width           =   1515
      End
      Begin VB.Label Label7 
         Caption         =   "País de Procedencia "
         Height          =   255
         Left            =   360
         TabIndex        =   33
         Top             =   1830
         Width           =   1635
      End
      Begin VB.Label Label8 
         Caption         =   "Representante "
         Height          =   195
         Left            =   360
         TabIndex        =   32
         Top             =   2850
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Localidad "
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   31
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "Cargo del Reprentante "
         Height          =   315
         Left            =   360
         TabIndex        =   30
         Top             =   3210
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "Fax "
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2550
         Width           =   1515
      End
      Begin VB.Label Label12 
         Caption         =   "Telf. Representante "
         Height          =   195
         Left            =   360
         TabIndex        =   28
         Top             =   3690
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   7335
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3840
         Picture         =   "FrmProvee.frx":1163
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton Transf 
         Caption         =   "Transf. a  Contab."
         Height          =   825
         Left            =   5955
         Picture         =   "FrmProvee.frx":15A5
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   195
         Width           =   825
      End
      Begin VB.CommandButton CmdIng 
         Caption         =   "&Ingreso"
         Height          =   675
         Left            =   660
         Picture         =   "FrmProvee.frx":19E7
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   270
         Width           =   775
      End
      Begin VB.CommandButton CmdModi 
         Caption         =   "&Modificar"
         Height          =   675
         Left            =   1740
         Picture         =   "FrmProvee.frx":1E29
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   270
         Width           =   780
      End
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Left            =   4860
         Picture         =   "FrmProvee.frx":226B
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   255
         Width           =   775
      End
      Begin VB.CommandButton CmdEli 
         Caption         =   "&Eliminar"
         Height          =   675
         Left            =   2820
         Picture         =   "FrmProvee.frx":26AD
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   255
         Width           =   775
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   7320
      Begin VB.ComboBox CmbOrden 
         Height          =   315
         ItemData        =   "FrmProvee.frx":2AEF
         Left            =   5265
         List            =   "FrmProvee.frx":2AF9
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   225
         Width           =   1575
      End
      Begin VB.TextBox TxFiltro 
         Height          =   300
         Left            =   1200
         TabIndex        =   15
         Text            =   "TxFiltro"
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "Orden :"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "Buscar   :"
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FrmProvee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adodc1 As ADODB.Recordset
Dim cSel1 As ADODB.Recordset
Dim cSel2 As ADODB.Recordset
Dim cSql1 As String, CSQL2 As String, cBase As String
Dim cSql3 As String, nT As Integer
Dim cCod As String, cDes As String
Dim nCom As Integer, nExiste As Integer
Dim nTra2 As Integer, nCursor As Integer
Dim nTra As Integer
Private Sub OculObj01(nTipo As Boolean)
Frame5.Visible = nTipo
Frame1.Visible = Not nTipo
Frame2.Visible = nTipo
Frame3.Visible = Not nTipo
DataGrid1.Visible = nTipo
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkAsignarRUC_Click()
    If chkAsignarRUC.Value Then
        Text1(1) = Text1(0)
    Else
        Text1(1) = ""
    End If
End Sub

Private Sub chkAsignarRUC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        chkAsignarRUC_Click
        SendKeys "{tab}"
    End If
End Sub

Private Sub CmbOrden_Click()            ' Ordenar por
nCom = CmbOrden.ListIndex
Set adodc1 = New ADODB.Recordset
Select Case nCom
Case 0
    adodc1.Open "Select PRVCCODIGO,PRVCNOMBRE,PRVCDIRECC,PRVCTELEF1 FROM MAEPROV ORDER BY PRVCCODIGO", VGCNx, adOpenStatic
Case 1
    adodc1.Open "Select PRVCCODIGO,PRVCNOMBRE,PRVCDIRECC,PRVCTELEF1 FROM MAEPROV ORDER BY PRVCNOMBRE", VGCNx, adOpenStatic
End Select
TxFiltro = ""
Set DataGrid1.DataSource = adodc1
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub
Private Sub CmdEli_Click()              ' Elimina
Dim nPosi As Integer
On Error GoTo EliErr

chkAsignarRUC.Enabled = False
If adodc1.RecordCount > 0 Then
    If MsgBox("Desea Eliminar Datos ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
        If Existe(1, adodc1("PRVCCODIGO"), "MovAlmcab", "Cacodpro", False) Then
            MsgBox "No se puede eliminar el Proveedor, porque tiene documentos Anexados", vbInformation, "Información"
            Exit Sub
        Else
            cBase = cRuta4
            If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
            Dim Nombre As String
               If Not adodc1.EOF Then
                  Nombre = "ANEXOPROV"
                  cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
                  Set cSel1 = New ADODB.Recordset
                  cSel1.Open cSql1, VGcnxCT, adOpenStatic
                  If Not cSel1.EOF Then
                     cAnexo = cSel1("ConcGral_Contec")
                     cSql1 = "DELETE FROM ANEXO Where TIPOANEX_CODIGO = '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("PRVCCODIGO")) & "'"
                     VGcnxCT.Execute cSql1
                  End If
                End If
                cSel1.Close
             End If
             
            cSql1 = "Delete from MaeProv where  PRVCCODIGO= '" & adodc1("PRVCCODIGO") & "'"
            nPosi = Pos_Dato(adodc1)
            nTra = 1
            VGCNx.BeginTrans
            VGCNx.Execute cSql1
            VGCNx.CommitTrans
            
            
            
            nTra = 0
            adodc1.Requery
            If nPosi <> 0 Then adodc1.AbsolutePosition = nPosi
        End If
    End If
    If DataGrid1.Visible Then DataGrid1.SetFocus
Else
    MsgBox "No existe registros para Eliminar", vbInformation, "Mensaje"
    Exit Sub
End If
Exit Sub
EliErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdGrabar_Click()           ' Grabar
Dim cMon As String
On Error GoTo GrabErr

If nT = 1 Then
  If Text1(0) <> "" Then
    If Existe(1, Text1(0), "MAEPROV", "PRVCCODIGO", False) Then
        MsgBox "El código de Proveedor, ya existe", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
    End If
  Else
        MsgBox "Ingrese código de Proveedor", vbInformation, "Mensaje"
        Text1(0).SetFocus: Exit Sub
  End If
End If
    
If Trim(Text1(2)) = "" Then
    MsgBox "Ingrese Nombre de Proveedor", vbInformation, "Mensaje"
    Text1(2).SetFocus: Exit Sub
End If
If Trim(Text1(3)) = "" Then
    MsgBox "Ingrese Dirección de Proveedor", vbInformation, "Mensaje"
    Text1(3).SetFocus: Exit Sub
End If
If Text1(1) <> "" Then
   If Validar_RUC(Text1(1)) = False Then
      Text1(1).SetFocus: Exit Sub
   End If
End If

If MsgBox("Es correcta la Información ?", vbQuestion + vbOKCancel, "Mensaje") = vbOK Then
    
    If nT = 1 Then      'Ingreso
        CSQL2 = "Insert Into MaeProv (PRVCCODIGO,PRVCRUC,PRVCNOMBRE,PRVCDIRECC,PRVCLOCALI,PRVCPAISAC,"
        CSQL2 = CSQL2 & "PRVCGIROAC,PRVCTELEF1,PRVCFAXACR,PRVCREPRES,PRVCCARREP,PRVCTELREP) VALUES "
        CSQL2 = CSQL2 & "('" & Text1(0) & "','" & Text1(1) & "','" & SupCadSQL(Text1(2)) & "','" & SupCadSQL(Text1(3)) & "',"
        CSQL2 = CSQL2 & "'" & Text1(4) & "','" & Text1(5) & "','" & Text1(6) & "','" & Text1(7) & "',"
        CSQL2 = CSQL2 & "'" & Text1(8) & "','" & SupCadSQL(Text1(9)) & "','" & Text1(10) & "','" & Text1(11) & "')"
        cCod = Text1(0)
        
    ElseIf nT = 2 Then     'Modificar
        CSQL2 = "Update MaeProv Set PRVCCODIGO='" & Text1(0) & "',PRVCRUC='" & Text1(1) & "',"
        CSQL2 = CSQL2 & "PRVCNOMBRE='" & SupCadSQL(Text1(2)) & "',PRVCDIRECC='" & SupCadSQL(Text1(3)) & "',PRVCLOCALI='" & Text1(4) & "',PRVCPAISAC='" & Text1(5) & "',"
        CSQL2 = CSQL2 & "PRVCGIROAC='" & Text1(6) & "',PRVCTELEF1='" & Text1(7) & "',"
        CSQL2 = CSQL2 & "PRVCFAXACR='" & Text1(8) & "',PRVCREPRES='" & SupCadSQL(Text1(9)) & "',"
        CSQL2 = CSQL2 & "PRVCCARREP='" & Text1(10) & "',PRVCTELREP='" & Text1(11) & "' "
        CSQL2 = CSQL2 & "Where PRVCCODIGO= '" & Trim(Text1(0)) & "'"
        cCod = Text1(0)
    End If
    
    nTra = 1
    VGCNx.BeginTrans
    VGCNx.Execute CSQL2
    VGCNx.CommitTrans
    nTra = 0
    
    Dim Nombre As String
    cBase = cRuta4
    If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
      'Se hace un enlace con los archivos de contabilidad, se busca y se graba
       If Not adodc1.EOF Then
          Nombre = "ANEXOPROV"
          cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
          Set cSel1 = New ADODB.Recordset
          cSel1.Open cSql1, VGcnxCT, adOpenStatic
          If Not cSel1.EOF Then
             cAnexo = cSel1("ConcGral_Contec")
          End If
          cSel1.Close
          cSql1 = "Select PRVCCODIGO,PRVCNOMBRE,PRVCRUC,PRVCDIRECC,PRVCTELEF1,PRVCREPRES FROM MAEPROV Where PRVCCODIGO='" & Trim(adodc1("PRVCCODIGO")) & "'"
          Set cSel2 = New ADODB.Recordset
          cSel2.Open cSql1, VGCNx, adOpenStatic
          If Not cSel2.EOF Then
            cSql1 = "Select * from ANEXO Where TIPOANEX_CODIGO= '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("PRVCCODIGO")) & "'"
            Set cSel1 = New ADODB.Recordset
            cSel1.Open cSql1, VGcnxCT, adOpenStatic
            If cSel1.RecordCount = 0 Then
               cSql3 = "Insert Into ANEXO (TIPOANEX_CODIGO,ANEX_CODIGO,ANEX_DESCRIPCION,ANEX_RUC,ANEX_DIRECCION,"
               cSql3 = cSql3 & "ANEX_TELEFONO,ANEX_REPRESENTANTE) values ('" & cAnexo & "','" & Trim(adodc1("PRVCCODIGO")) & "','" & IIf(Trim(cSel2("PRVCNOMBRE")) <> "", Trim(Mid(cSel2("PRVCNOMBRE"), 1, 50)), "0") & "','" & IIf(Trim(cSel2("PRVCRUC")) <> "", cSel2("PRVCRUC"), "0") & "',"
               cSql3 = cSql3 & "'" & IIf(Trim(cSel2("PRVCDIRECC")) <> "", Mid(cSel2("PRVCDIRECC"), 1, 50), "0") & "','" & IIf(Trim(cSel2("PRVCTELEF1")) <> "", Mid(cSel2("PRVCTELEF1"), 1, 15), "0") & "','" & IIf(Trim(cSel2("PRVCREPRES")) <> "", cSel2("PRVCREPRES"), "0") & "')"
               VGcnxCT.Execute cSql3
            Else
                cSql3 = "Update ANEXO Set TIPOANEX_CODIGO ='" & cAnexo & "' ,ANEX_CODIGO='" & Trim(adodc1("PRVCCODIGO")) & "' ,"
                cSql3 = cSql3 & "ANEX_DESCRIPCION='" & IIf(Trim(cSel2("PRVCNOMBRE")) <> "", Trim(Mid(cSel2("PRVCNOMBRE"), 1, 50)), "0") & "',ANEX_RUC='" & IIf(cSel2("PRVCRUC") <> "", cSel2("PRVCRUC"), "0") & "',ANEX_DIRECCION='" & IIf(Trim(cSel2("PRVCDIRECC")) <> "", Mid(cSel2("PRVCDIRECC"), 1, 50), "0") & "',"
                cSql3 = cSql3 & "ANEX_TELEFONO='" & IIf(Trim(cSel2("PRVCTELEF1")) <> "", Mid(cSel2("PRVCTELEF1"), 1, 15), "0") & "',ANEX_REPRESENTANTE='" & IIf(Trim(cSel2("PRVCREPRES")) <> "", Trim(cSel2("PRVCREPRES")), "0") & "' Where TIPOANEX_CODIGO = '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("PRVCCODIGO")) & "'"
                VGcnxCT.Execute cSql3
            End If
            cSel1.Close
         End If
         cSel2.Close
      End If
    End If

    adodc1.Requery
    
    adodc1.Find "PRVCCODIGO = '" & cCod & "'"
End If


If nT = 1 Then
    Limpiar
    Text1(0).SetFocus
Else
    CmdSalir2_Click
End If
Exit Sub
GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then VGCNx.RollbackTrans
    If nTra2 = 1 Then VGCNx.RollbackTrans
End Sub

Private Sub CmdIng_Click()      'Ingresar
nT = 1
Frame3.Caption = "Ingreso de Datos del Proveedor"
OculObj01 (False)
Limpiar
Text1(0).Enabled = True
Text1(0).SetFocus
chkAsignarRUC.Value = 1
chkAsignarRUC.Enabled = True
End Sub

Private Sub CmdModi_Click()     'Modificar
chkAsignarRUC.Enabled = False
If adodc1.RecordCount > 0 Then
    nT = 2
    Frame3.Caption = "Modificación de Datos de Proveedor"
    OculObj01 (False)
    Limpiar
    If Not adodc1.EOF Then
       If Not IsNull(adodc1("PRVCCODIGO")) Or adodc1("PRVCCODIGO") <> "" Then cCod = adodc1("PRVCCODIGO")
    End If
    Mostrar (cCod)
    Text1(0).Enabled = False
    If Text1(1).Visible Then Text1(1).SetFocus
Else
    MsgBox "No existen registros", vbInformation, "Mensaje"
End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdSalir2_Click()   'Salida de la segunda pantalla
OculObj01 (True)
DataGrid1.SetFocus
End Sub

Private Sub Command1_Click()
    Dim CADENA As String
    'Dim cFormato As String
    'Dim cDireccion As String
    'Dim cRuc As String
    Dim cNomRepor  As String
    Dim aBusca As New ADODB.Recordset

cNomRepor = "proveedor.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Listado de Proveedores"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
   
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If
End Sub

Private Sub CmdImprimir_Click()
Dim CADENA As String
Dim cNomRepor  As String

cNomRepor = "proveedor.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Proveedores"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    
    CrystalReport1.formulas(1) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If

End Sub


Private Sub Form_Activate()
TxFiltro = ""
CmbOrden.ListIndex = 0
If DataGrid1.Visible Then DataGrid1.SetFocus
End Sub
Private Sub Form_Load()
Dim RUTA As String
Dim NAMEBD As String
central Me          ' Centrar Formulario
Init_ControlDataGrid DataGrid1

Limpiar
OculObj01 (True)
Set adodc1 = New ADODB.Recordset
adodc1.Open "Select PRVCCODIGO,PRVCNOMBRE,PRVCDIRECC,PRVCTELEF1 FROM MAEPROV ORDER BY PRVCCODIGO", VGCNx, adOpenStatic
Set DataGrid1.DataSource = adodc1
DataGrid1.Refresh
End Sub
Private Sub Limpiar()   'Limpia variables
Dim n As Integer
For n = 0 To 11: Text1(n) = "": Next
End Sub

Private Sub Text1_DblClick(Index As Integer)
If Index = 6 Then
        Dim Adodc3 As ADODB.Recordset
        Set Adodc3 = New ADODB.Recordset
        Adodc3.Open "SELECT TCLAVE,TDESCRI FROM TabAYU where  TCOD = '62'", VGCNx, adOpenStatic, adLockOptimistic
        frmReferencia.Conectar Adodc3, "SELECT TCLAVE,TDESCRI FROM TabAYU where  TCOD = '62'"
        frmReferencia.Label1.Caption = "Giro del proveedor"
        frmReferencia.Show vbModal
        Adodc3.Close
        If vGUtil(1) <> "" Then
                Text1(6) = vGUtil(1)
        End If
End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Enfoque Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        If Trim(Text1(0)) <> "" Then
            If Existe(1, Text1(0), "MAEPROV", "PRVCCODIGO", False) Then
                MsgBox "El código de Proveedor, ya existe", vbInformation, "Mensaje"
                Text1(0).SetFocus
                Exit Sub
            End If
            'chkAsignarRUC_Click
            SendKeys "{tab}"
            'Text1(1).SetFocus:
            Exit Sub
        Else
            MsgBox "Ingrese código de Proveedor", vbInformation, "Mensaje"
            Text1(0).SetFocus: Exit Sub
        End If
    ElseIf Index = 1 Then
        If Validar_RUC(Text1(1)) = False Then
            Text1(1).SetFocus
            Exit Sub
        Else
            SendKeys "{tab}"
        End If
    ElseIf Index = 2 Then
        If Trim(Text1(2)) <> "" Then
            Text1(3).SetFocus: Exit Sub
        Else
            MsgBox "Ingrese Nombre de Proveedor", vbInformation, "Mensaje"
            Text1(2).SetFocus: Exit Sub
        End If
    ElseIf Index = 3 Then
        If Trim(Text1(3)) <> "" Then
            Text1(4).SetFocus: Exit Sub
        Else
            MsgBox "Ingrese Dirección de Proveedor", vbInformation, "Mensaje"
            Text1(3).SetFocus: Exit Sub
        End If
    Else
        If Index <> 11 Then
           If Index = 1 Then
              If Text1(1) <> "" Then
                 If Validar_RUC(Text1(1)) = False Then
                    Text1(1).SetFocus: Exit Sub
                 Else
                    Text1(2).SetFocus
                 End If
              End If
           End If
           SendKeys "{tab}"
        Else
           Cmdgrabar.SetFocus
        End If
    End If
    
Else
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    If Index = 1 Then
    If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
    End If
End If
End Sub

Private Sub Transf_Click()
Dim Nombre As String
cBase = cRuta4
 If UCase(Dir$(cBase)) = VGNameCont & ".MDB" Then
    'Se hace un enlace con los archivos de contabilidad, se busca y se graba
   If Not adodc1.EOF Then
      adodc1.MoveFirst
      Nombre = "ANEXOPROV"
      cSql1 = "Select ConcGral_Contec from Conceptos_Generales Where ConcGral_Codigo= '" & UCase(Nombre) & "'"
      Set cSel1 = New ADODB.Recordset
      cSel1.Open cSql1, VGcnxCT, adOpenStatic
      If Not cSel1.EOF Then
         cAnexo = cSel1("ConcGral_Contec")
      End If
      cSel1.Close
      
         Do While Not adodc1.EOF
            cSql1 = "Select * from ANEXO Where TIPOANEX_CODIGO= '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("PRVCCODIGO")) & "'"
            Set cSel1 = New ADODB.Recordset
            cSel1.Open cSql1, VGcnxCT, adOpenStatic
            
            cSql1 = "Select PRVCCODIGO,PRVCNOMBRE,PRVCRUC,PRVCDIRECC,PRVCTELEF1,PRVCREPRES FROM MAEPROV Where PRVCCODIGO='" & Trim(adodc1("PRVCCODIGO")) & "'"
            Set cSel2 = New ADODB.Recordset
            cSel2.Open cSql1, VGCNx, adOpenStatic
            If Not cSel2.EOF Then
               If cSel1.RecordCount = 0 Then
                  cSql3 = "Insert Into ANEXO (TIPOANEX_CODIGO,ANEX_CODIGO,ANEX_DESCRIPCION,ANEX_RUC,ANEX_DIRECCION,"
                  cSql3 = cSql3 & "ANEX_TELEFONO,ANEX_REPRESENTANTE) values ('" & cAnexo & "','" & Trim(adodc1("PRVCCODIGO")) & "','" & IIf(Trim(cSel2("PRVCNOMBRE")) <> "", Trim(Mid(cSel2("PRVCNOMBRE"), 1, 50)), "0") & "','" & IIf(Trim(cSel2("PRVCRUC")) <> "", cSel2("PRVCRUC"), "0") & "',"
                  cSql3 = cSql3 & "'" & IIf(Trim(cSel2("PRVCDIRECC")) <> "", Mid(cSel2("PRVCDIRECC"), 1, 50), "0") & "','" & IIf(Trim(cSel2("PRVCTELEF1")) <> "", Mid(cSel2("PRVCTELEF1"), 1, 15), "0") & "','" & IIf(Trim(cSel2("PRVCREPRES")) <> "", cSel2("PRVCREPRES"), "0") & "')"
                  VGcnxCT.Execute cSql3
               Else
                   cSql3 = "Update ANEXO Set TIPOANEX_CODIGO ='" & cAnexo & "' ,ANEX_CODIGO='" & Trim(adodc1("PRVCCODIGO")) & "' ,"
                   cSql3 = cSql3 & "ANEX_DESCRIPCION='" & IIf(Trim(cSel2("PRVCNOMBRE")) <> "", Trim(Mid(cSel2("PRVCNOMBRE"), 1, 50)), "0") & "',ANEX_RUC='" & IIf(cSel2("PRVCRUC") <> "", cSel2("PRVCRUC"), "0") & "',ANEX_DIRECCION='" & IIf(Trim(cSel2("PRVCDIRECC")) <> "", Mid(cSel2("PRVCDIRECC"), 1, 50), "0") & "',"
                   cSql3 = cSql3 & "ANEX_TELEFONO='" & IIf(Trim(cSel2("PRVCTELEF1")) <> "", Mid(cSel2("PRVCTELEF1"), 1, 15), "0") & "',ANEX_REPRESENTANTE='" & IIf(Trim(cSel2("PRVCREPRES")) <> "", Trim(cSel2("PRVCREPRES")), "0") & "' Where TIPOANEX_CODIGO = '" & cAnexo & "' and ANEX_CODIGO = '" & Trim(adodc1("PRVCCODIGO")) & "'"
                   VGcnxCT.Execute cSql3
               End If
            End If
            cSel1.Close
            cSel2.Close
            adodc1.MoveNext
         Loop
       End If
      adodc1.MoveFirst
 End If
End Sub

Private Sub TxFiltro_Change()
If adodc1.RecordCount > 0 Then
    If Trim(TxFiltro) <> "" And TxFiltro.Visible Then
        nCursor = adodc1.Bookmark
        adodc1.AbsolutePosition = 1
        adodc1.MoveFirst
        
        Select Case CmbOrden.ListIndex
        Case 0
            adodc1.Find "PRVCCODIGO LIKE '" & Trim(UCase(TxFiltro)) & "*'"
        Case 1
            adodc1.Find "PRVCNOMBRE LIKE '" & Trim(UCase(TxFiltro)) & "*' "
        End Select
        If adodc1.EOF Then adodc1.AbsolutePosition = nCursor
    End If
End If
End Sub

Private Sub Mostrar(cC1 As String) 'Muestra los datos
Dim cSqlM As String, cSelM As ADODB.Recordset
If Trim(cC1) = "" Then
    MsgBox "No hay registros para mostrar", vbInformation, "Mensaje"
    Exit Sub
End If

cSqlM = "Select * From MaeProv Where PRVCCODIGO= '" & cC1 & "'"
Set cSelM = New ADODB.Recordset
cSelM.Open cSqlM, VGCNx, adOpenStatic
If cSelM.RecordCount > 0 Then
    Text1(0) = cSelM("PRVCCODIGO")
    If Not IsNull(cSelM("PRVCRUC")) Then Text1(1) = cSelM("PRVCRUC")
    If Not IsNull(cSelM("PRVCNOMBRE")) Then Text1(2) = cSelM("PRVCNOMBRE")
    If Not IsNull(cSelM("PRVCDIRECC")) Then Text1(3) = cSelM("PRVCDIRECC")
    If Not IsNull(cSelM("PRVCLOCALI")) Then Text1(4) = cSelM("PRVCLOCALI")
    If Not IsNull(cSelM("PRVCPAISAC")) Then Text1(5) = cSelM("PRVCPAISAC")
    If Not IsNull(cSelM("PRVCGIROAC")) Then Text1(6) = cSelM("PRVCGIROAC")
    If Not IsNull(cSelM("PRVCTELEF1")) Then Text1(7) = cSelM("PRVCTELEF1")
    If Not IsNull(cSelM("PRVCFAXACR")) Then Text1(8) = cSelM("PRVCFAXACR")
    If Not IsNull(cSelM("PRVCREPRES")) Then Text1(9) = cSelM("PRVCREPRES")
    If Not IsNull(cSelM("PRVCCARREP")) Then Text1(10) = cSelM("PRVCCARREP")
    If Not IsNull(cSelM("PRVCTELREP")) Then Text1(11) = cSelM("PRVCTELREP")
Else
    MsgBox "No existe registro", vbInformation, "Mensaje"
    CmdSalir2_Click
End If
cSelM.Close
End Sub
