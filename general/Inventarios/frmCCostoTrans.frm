VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form FrmCCostoTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centro de Costos"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6825
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   240
      TabIndex        =   12
      Top             =   2760
      Width           =   6375
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   675
         Index           =   3
         Left            =   3240
         Picture         =   "frmCCostoTrans.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   675
         Index           =   2
         Left            =   2280
         Picture         =   "frmCCostoTrans.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   675
         Index           =   4
         Left            =   4200
         Picture         =   "frmCCostoTrans.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   675
         Index           =   0
         Left            =   360
         Picture         =   "frmCCostoTrans.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Grabar"
         Height          =   675
         Index           =   1
         Left            =   1320
         Picture         =   "frmCCostoTrans.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   675
         Index           =   5
         Left            =   5160
         Picture         =   "frmCCostoTrans.frx":154A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   360
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Codigo"
         Height          =   675
         Index           =   7
         Left            =   360
         Picture         =   "frmCCostoTrans.frx":198C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Transf"
         Height          =   675
         Index           =   6
         Left            =   4200
         Picture         =   "frmCCostoTrans.frx":1DCE
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   775
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Lista de Centros de Costos:"
      Height          =   2535
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   6345
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1935
         Left            =   360
         TabIndex        =   11
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3413
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
   End
   Begin VB.Frame FrameU 
      Caption         =   "Modificación  de C. de Costo"
      Height          =   2295
      Index           =   2
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6360
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   690
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo de Centro de Costo :"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción:"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   3
         Top             =   1125
         Width           =   2775
      End
   End
   Begin VB.Frame FrameU 
      Caption         =   "Ingreso de Centro de Costos :"
      Height          =   2295
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   6360
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   2640
         MaxLength       =   6
         TabIndex        =   6
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   2640
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   9
         Top             =   1125
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo de Centro de Costo :"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FrmCCostoTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Adoreg1 As ADODB.Recordset
Dim RegActual As Integer
Public PN_TOTAL As Double
Dim nFra As Integer

Private Sub cmdBotones_Click(Index As Integer)
Dim flag As Boolean
Dim cTablaSource As String, cTablaDestination As String
Dim src As String
Dim ncar As String
Dim cSel2 As New ADODB.Recordset
Dim criterio As String

Select Case Index

 Case 0: 'Nuevo
         FrameU(nFra).Visible = False
         FrameU(1).Visible = True
         nFra = 1
         
         Dim otext As TextBox
         For Each otext In Me.Text1
          otext.text = ""
         Next
         Text1(0).SetFocus
         Botones_Set False
         
 Case 1: 'Grabar
 
       If FrameU(1).Visible Then 'Nuevo
          flag = False
          
            If Text1(0).text = "" Then
              MsgBox "Ud. No ha ingresado el Código de Centro de Costo", vbInformation, "Ingreso de Datos"
              Text1(0).SetFocus
              Exit Sub
            End If
            
          If Text1(1).text = "" Then
            MsgBox "Ud. No ha ingresado la Descripción", vbInformation, "Ingreso de Datos"
            Text1(1).SetFocus
            Exit Sub
          End If
          
          'buscar igual codigo
          If Not Adoreg1.EOF Then
                
              With Adoreg1
                 .MoveFirst
                 
                 Do While Not .EOF
                    If UCase(Text1(0).text) = .Fields("CENCOST_CODIGO") Then
                     flag = True
                     Text1(0).text = ""
                     MsgBox "El Centro de Costo Ya Existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                     Exit Do
                    End If
                    .MoveNext
                Loop
                 
              End With
           End If
           
          If Not flag Then
            'pasa
            Adoreg1.AddNew
            Adoreg1.Fields("CENCOST_CODIGO") = IIf(IsNull(Text1(0).text), "", Text1(0).text)
            Adoreg1.Fields("CENCOST_DESCRIPCION") = IIf(IsNull(Text1(1).text), "", Text1(1).text)
            Adoreg1.UpdateBatch
            Adoreg1.Requery
            DataGrid1.Refresh
            'Set DataGrid1.DataSource = Adoreg1
            FrameU(nFra).Visible = False
            FrameU(0).Visible = True
            nFra = 0
            Botones_Set True
          End If
         End If
         
         If FrameU(2).Visible Then 'Editar
            
            If Text2(1).text = "" Then
              MsgBox "Ud. No ha ingresado la Descripción", vbInformation, "Modificación de Datos"
              Text2(1).SetFocus
              Exit Sub
            End If
          
            Adoreg1.Fields("CENCOST_CODIGO") = IIf(IsNull(Text2(0).text), "", Text2(0).text)
            Adoreg1.Fields("CENCOST_DESCRIPCION") = IIf(IsNull(Text2(1).text), "", Text2(1).text)
            Adoreg1.UpdateBatch
            Adoreg1.Requery
          
           FrameU(nFra).Visible = False
           FrameU(0).Visible = True
           nFra = 0
           Botones_Set True
         End If
         
         SetDataGrid
         
 Case 2: 'Editar
 
         If Adoreg1.Bookmark Then
                FrameU(nFra).Visible = False
                FrameU(2).Visible = True
                nFra = 2
                Text2(0).text = IIf(IsNull(Adoreg1.Fields("CENCOST_CODIGO")), "", Adoreg1.Fields("CENCOST_CODIGO"))
                Text2(1).text = IIf(IsNull(Adoreg1.Fields("CENCOST_DESCRIPCION")), "", Adoreg1.Fields("CENCOST_DESCRIPCION"))
                Text2(1).SetFocus
                Botones_Set False
         Else
                MsgBox "Debe seleccionar un Registro para editarlo", vbInformation
                Botones_Set False
                cmdBotones_Click 5
         End If
         
 Case 3: 'Eliminar
 
    Dim op As Integer
    Dim cSel1 As New ADODB.Recordset
    Dim csql As String
    Dim PASE As Boolean
    
   PASE = False
   
    If Adoreg1.RecordCount <> 0 Then
        Set cSel1 = VGCNx.Execute("SELECT * from CENTRO_COSTOS WHERE CENCOST_CODIGO='" & Text1(0).text & "'")
        If cSel1.RecordCount() > 0 Then
           PASE = True
         'Adoreg1.MovePrevious
        End If
         op = MsgBox("Seguro de Eliminar el Registro ", vbYesNo, "Eliminación de Registro")
          If op = vbYes Then
             Set cSel1 = Nothing
             Set cSel1 = VGCNx.Execute("DELETE CENTRO_COSTOS WHERE CENCOST_CODIGO='" & Text1(0).text & "'")
             If Adoreg1.RecordCount = 0 Then
                    Botones_Init True
             Else
                     Botones_Set True
             End If
             
          End If
         DataGrid1.Refresh
   
       
   End If
   
 Case 4: 'Imprimir

Dim cadena As String
Dim cNomRepor  As String

cNomRepor = "centrocosto.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Centro de Costo"
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



 
 
 Case 5: 'Salir , Cancelar
 
         If cmdBotones(5).Caption = "&Salir" Then
             Unload Me
         Else
            cmdBotones(5).Caption = "&Salir"
            FrameU(nFra).Visible = False
            FrameU(0).Visible = True
            nFra = 0
                If Adoreg1.RecordCount = 0 Then
                 Botones_Init True
                Else
                 Botones_Set True
                End If
            If cmdBotones(7).Visible Then cmdBotones(7).Visible = False
            'If cmdBotones(8).Visible Then cmdBotones(8).Visible = False
          
         End If
 
 Case 6: 'Transferencia
  
'   If Adoreg1.RecordCount <> 0 Then    'Transferencia Clase 9
'
'      If Len(Trim(Adoreg1.Fields("cencost_codigo"))) = 6 Then
'               frmTransCCosto.Caption = "TABLA DE TRANSFERENCIA DE COSTO"
'               frmTransCCosto.DataGrid1.Caption = "Costo  :  " & Adoreg1.Fields("CENCOST_CODIGO") & Space(5) & Adoreg1.Fields("CENCOST_DESCRIPCION")
'               frmTransCCosto.show 1
'
'       Else
'          MsgBox ("La Distribución de los Costos se hace a nivel de más detalle"), vbExclamation + vbOKOnly, "Advertencia"
'          Exit Sub
'      End If
' End If

 Case 7: 'Buscar por Codigo
         
          'Dim criterio As String
          If Adoreg1.RecordCount <> 0 Then
           Adoreg1.MoveFirst
           src = InputBox$("Ingrese Código", "Búsqueda")
           ncar = Str$(Len(src))
           criterio = "Left(CCODCLI," & ncar & ") = '" & src & "'"
           criterio = "CENCOST_CODIGO LIKE '" & src & "%'"
          
           Adoreg1.Find criterio
           If Adoreg1.EOF Then
              MsgBox "No se encontro el registro", vbExclamation + vbOKOnly, "Advertencia"
              Adoreg1.MoveFirst
           End If
         DataGrid1.Refresh
         DataGrid1.SetFocus
        End If
         
 
' Case 8: 'Buscar por Descripción
'         If Adoreg1.RecordCount <> 0 Then
'            Adoreg1.MoveFirst
'          SRC = InputBox$("Ingrese Descripción", "Búsqueda")
'          ncar = Str$(Len(SRC))
'          criterio = "MID$(CNOMCLI,1," + ncar + ") = " & Chr$(34) + SRC + Chr$(34)
'          criterio = "CENCOST_DESCRIPCION LIKE '" & SRC & "*'"
'           Adoreg1.Find criterio
'           If Adoreg1.EOF Then
'               MsgBox "No se encontro el registro", vbExclamation + vbOKOnly, "Advertencia"
'               Adoreg1.MoveFirst
'           End If
'          DataGrid1.Refresh
'         DataGrid1.SetFocus
'        End If
End Select
End Sub

Public Sub Botones_Set(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 cmdBotones(0).Enabled = flag 'Nuevo
 cmdBotones(3).Enabled = flag 'Eliminar
 cmdBotones(1).Enabled = Not flag 'Grabar
 cmdBotones(2).Enabled = flag 'Editar
 cmdBotones(4).Enabled = flag 'Buscar
 cmdBotones(6).Enabled = flag 'Transferencia
 If flag Then
  cmdBotones(5).Caption = "&Salir" 'Salir
 Else
  cmdBotones(5).Caption = "&Cancelar"
 End If
End Sub
Public Sub Botones_Init(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 cmdBotones(0).Enabled = flag 'Nuevo
 cmdBotones(3).Enabled = Not flag 'Eliminar
 cmdBotones(1).Enabled = Not flag 'Grabar
 cmdBotones(2).Enabled = Not flag 'Editar
 cmdBotones(4).Enabled = Not flag 'Buscar
 cmdBotones(6).Enabled = Not flag 'Transferencia
 cmdBotones(5).Caption = "&Salir" 'Salir
End Sub

Private Sub DataGrid1_Click()
 RegActual = IIf(IsNull(DataGrid1.Bookmark), 0, DataGrid1.Bookmark)
End Sub

Private Sub Form_Activate()
DataGrid1.SetFocus

End Sub

Private Sub Form_Load()
 Dim fra As Frame
 central Me

    ADOConectar
    Init_ControlDataGrid DataGrid1
    If Adoreg1.RecordCount = 0 Then
         Botones_Init True
    Else
          Botones_Set True
    End If
    SetDataGrid

 For Each fra In Me.FrameU
  fra.Visible = False
 Next
 FrameU(0).Visible = True
 nFra = 0
 'Combo1(0).ListIndex = 0
 PN_TOTAL = 0
 DataGrid1.Refresh
End Sub

Public Sub ADOConectar()

   Set Adoreg1 = New ADODB.Recordset
   Adoreg1.Open "SELECT cencost_codigo,cencost_descripcion FROM centro_costos  order by cencost_codigo ", VGCNx, adOpenStatic, adLockOptimistic
   Set DataGrid1.DataSource = Adoreg1
End Sub

Public Sub SetDataGrid()
 DataGrid1.Refresh
 
 DataGrid1.Columns(0).Caption = "Código"
 DataGrid1.Columns(1).Caption = "Descripción"
 
 DataGrid1.Columns(0).DataField = "CENCOST_CODIGO"
 DataGrid1.Columns(1).DataField = "CENCOST_DESCRIPCION"
 
 DataGrid1.Columns(0).Width = 1000
 DataGrid1.Columns(1).Width = 5000

DataGrid1.Refresh

End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0: Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then
   If KeyAscii = 13 Then
 
      If Text1(0) <> "" Then
            If Len(Text1(0)) = 2 Or Len(Text1(0)) = 4 Or Len(Text1(0)) = 6 Then
               Text1(1).SetFocus
            Else
             MsgBox ("Los codigos de Centros de Costos son de 2, 4 o 6 Digitos"), vbExclamation + vbOKOnly, "Advertencia"
             Text1(0) = ""
             Text1(0).SetFocus
             Exit Sub
            End If
      Else
        MsgBox ("Ingrese Codigo de Centro de Costo"), vbExclamation + vbOKOnly, "Advertencia"
        Text1(0).SetFocus
      End If
    
   Else
'     If (Chr$(KeyAscii) > "9" Or Chr$(KeyAscii) < "0") And KeyAscii <> 8 Then KeyAscii = 0
   End If
    
Else

  If KeyAscii = 13 Then
     cmdBotones(1).SetFocus
  Else
  
    If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
  End If
  
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim cSel1 As New ADODB.Recordset
Dim cSql1 As String
If Index = 0 Then

   If Text1(0) <> "" Then
        If Len(Text1(0)) = 2 Or Len(Text1(0)) = 4 Or Len(Text1(0)) = 6 Then
          cSql1 = "Select * from CENTRO_COSTOS WHERE CENCOST_CODIGO='" & Text1(0) & "'"
          cSel1.Open cSql1, VGCNx, adOpenDynamic, adLockOptimistic

          If cSel1.RecordCount = 0 Then
       '       ValidCCosto (Text1(0))
          Else
            MsgBox ("Codigo de Costo ya Existe"), vbExclamation + vbOKOnly, "Advertencia"
            cSel1.Close
            Text1(0) = ""
            Text1(0).SetFocus
            Exit Sub
          End If
          cSel1.Close
          
       Else
          MsgBox ("El Codigo de Centro de Costo puede tener 2,4 0 6 Digitos"), vbExclamation + vbOKOnly, "Advertencia"
          Text1(0) = ""
          Text1(0).SetFocus
       End If
    
    End If

End If
End Sub
Sub ValidCCosto(cCodigo As String)
Dim cSel1 As New ADODB.Recordset
Dim cSql1 As String

cSql1 = ""

If Len(cCodigo) = 4 Then
   
   cSql1 = "Select * from CENTRO_COSTOS WHERE CENCOST_CODIGO='" & Mid(cCodigo, 1, 2) & "'"
   cSel1.Open cSql1, VGCNx, adOpenDynamic, adLockOptimistic
   
   If cSel1.RecordCount = 0 Then
      MsgBox "Debe ingresar primero Centro de Costo de Nivel Superior", vbInformation + vbOKOnly, "Advertencia"
      Text1(0) = ""
      Text1(0).SetFocus
      Exit Sub
   End If
   cSel1.Close
   
ElseIf Len(cCodigo) = 6 Then

   cSql1 = "Select * from CENTRO_COSTOS WHERE CENCOST_CODIGO=LEFT('" & cCodigo & "',4) ORDER BY CENCOST_CODIGO"
  cSel1.Open cSql1, VGCNx, adOpenDynamic, adLockOptimistic
  If cSel1.RecordCount() = 0 Then
      MsgBox "Debe ingresar primero Centro de Costo de Nivel Superior", vbInformation + vbOKOnly, "Advertencia"
      Text1(0) = ""
      Text1(0).SetFocus
      Exit Sub
   End If
   cSel1.Close
End If
End Sub

Private Sub Text2_GotFocus(Index As Integer)
Text2(Index).SelStart = 0: Text2(Index).SelLength = Len(Text2(Index))
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then
    If KeyAscii = 13 Then
       SendKeys "{tab}"
       KeyAscii = 0
    End If
Else
    If KeyAscii = 13 Then
       cmdBotones(1).SetFocus
    Else
    
      If KeyAscii = 39 Or KeyAscii = 34 Then KeyAscii = 0
      KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
    End If
    
End If
End Sub

