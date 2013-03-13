VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmZonaVendedor 
   Caption         =   "Mantenimiento Zona - Vendedor"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   1050
      Left            =   1440
      TabIndex        =   8
      Top             =   5160
      Width           =   5655
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   0
         Left            =   225
         Picture         =   "FrmZonaVendedor.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   1320
         Picture         =   "FrmZonaVendedor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   2440
         Picture         =   "FrmZonaVendedor.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   4
         Left            =   4590
         Picture         =   "FrmZonaVendedor.frx":0CC6
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   3
         Left            =   3510
         Picture         =   "FrmZonaVendedor.frx":1108
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   180
         Width           =   870
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "FrmZonaVendedor.frx":154A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TDBGridProducto"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "FrmZonaVendedor.frx":1566
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cAcepta"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cCancela"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   3375
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   8625
         Begin VB.ComboBox cmbZona 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1200
            Width           =   3255
         End
         Begin VB.ComboBox cmbVendedor 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3960
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label lbl 
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   2160
            TabIndex        =   14
            Top             =   1320
            Width           =   1605
         End
         Begin VB.Label lbl 
            Caption         =   "Vendedor"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   2160
            TabIndex        =   6
            Top             =   600
            Width           =   1605
         End
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         HelpContextID   =   8
         Left            =   -70320
         TabIndex        =   3
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         HelpContextID   =   7
         Left            =   -72240
         TabIndex        =   2
         Top             =   4080
         Width           =   1335
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGridProducto 
         Height          =   4095
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   7223
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   13160660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2752"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2752"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   13160660
         RowDividerColor =   13160660
         RowSubDividerColor=   13160660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Arial"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Arial"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Arial"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "FrmZonaVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modoinsert, modoedit As Boolean
Dim i_filaorigen As Integer
Dim i_codigovendedor, i_codigozona As String
''''''''''''''''''''''''
Dim ArregloVendedor()
Dim ArregloZona()

Private Sub cAcepta_Click()
    
   Dim RS As New ADODB.Recordset
   Dim sql As String
   Dim J As Integer
 
   Dim s_codigovendedor As String
   Dim s_codigozona As String
   
   On Error GoTo CONTROLERRORES
   ''''''''
     If cmbVendedor.ListIndex <> -1 Then
        s_codigovendedor = ArregloVendedor(0, cmbVendedor.ListIndex)
     Else
        s_codigovendedor = ""
     End If
     If cmbZona.ListIndex <> -1 Then
        s_codigozona = ArregloZona(0, cmbZona.ListIndex)
     Else
        s_codigozona = ""
     End If
        
   If modoinsert = True Then
   
         If Validar_CodigosDuplicados("INSERT") = True Then
            MsgBox "Registro Duplicado", vbCritical, "Error"
            cAcepta.Enabled = False
            Exit Sub
          End If
               
          sql = "INSERT INTO vt_zonavendedor " & _
               "(vendedorcodigo,zonacodigo," & _
               "usuariocodigo,fechaact)" & _
               "VALUES " & _
               "('" & s_codigovendedor & "'," & _
               "'" & s_codigozona & "'," & _
               "'" & g_usuario & "','" & Date & "')"

          cn.Execute sql
                   
    ElseIf modoedit = True Then
   
             If Validar_CodigosDuplicados("UPDATE", i_filaorigen) = True Then
               MsgBox "Registro Duplicado", vbCritical, "Error"
               cAcepta.Enabled = False
               Exit Sub
             End If
                                 
            sql = "UPDATE vt_zonavendedor SET " & _
               "vendedorcodigo='" & s_codigovendedor & "'," & _
               "zonacodigo='" & s_codigozona & "'," & _
               "usuariocodigo='" & g_usuario & "'," & _
               "fechaact='" & Date & "' " & _
               "WHERE vendedorcodigo='" & i_codigovendedor & "' " & _
               "AND zonacodigo='" & i_codigozona & "'"
    
            cn.Execute sql
            
  End If
 '******************************************************************************************
        
 TDBGridProducto.Refresh
      
 Mostrar_Data
 MostrarOcultar_Botones (True)
 '''''''''
 modoinsert = False
 modoedit = False
 '''''''''
 SSTab1.TabEnabled(0) = True
 
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'cg.RollbackTrans
       Resume Next
    End If
       
End Sub

Private Sub cCancela_Click()
    SSTab1.TabEnabled(0) = True
    SSTab1.Tab = 0
    SSTab1.SetFocus
    MostrarOcultar_Botones (True)
    modoinsert = False
    modoedit = False
End Sub

Private Sub cmbzona_Click()
  cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbzona_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub

Private Sub cmbVendedor_Click()
  cAcepta.Enabled = Validar_DatosNulos()
End Sub

Private Sub cmbvendedor_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 13 Then SendKeys "{TAB}"
End Sub


Private Sub cmdBotones_Click(Index As Integer)
  Dim J As Integer
  Dim sql As String
  Dim OBJ As Object
  
  On Error GoTo CONTROLERRORES
  
  SSTab1.TabEnabled(1) = True
  '''''
  Select Case Index
     Case 0   'nuevo
            For Each OBJ In Me.Controls
                If TypeOf OBJ Is ComboBox Then
                    OBJ.ListIndex = -1
                End If
            Next
            SSTab1.Tab = 1
            modoinsert = True
            MostrarOcultar_Botones (False)
            cmbVendedor.SetFocus
        
     Case 1   'modificar
     
         If TDBGridProducto.Row < 0 Then
            Exit Sub
         End If
         
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(0).Text), cmbVendedor, ArregloVendedor)
             Call fncSeleccionaCombo(Trim(TDBGridProducto.Columns(2).Text), cmbZona, ArregloZona)
             
             i_codigovendedor = TDBGridProducto.Columns(0).Text
             i_codigozona = TDBGridProducto.Columns(2).Text
                              
        modoedit = True
        SSTab1.Tab = 1
        MostrarOcultar_Botones (False)
        i_filaorigen = TDBGridProducto.Row
        cmbVendedor.SetFocus
      
        '''''''''
      
     Case 2   'eliminar
     If TDBGridProducto.Row < 0 Then
            Exit Sub
     End If
         
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          sql = "DELETE FROM vt_zonavendedor WHERE " & _
                "vendedorcodigo = '" & TDBGridProducto.Columns(0).Text & _
                "' AND zonacodigo = '" & TDBGridProducto.Columns(2).Text & "'"
          cn.Execute sql
          Mostrar_Data
       End If
        
     Case 3   'imprimir
         'oCrystalReport.DiscardSavedData = True
         'oCrystalReport.Action = 1
     Case 4  ' salir
       Unload Me
  End Select
Exit Sub
CONTROLERRORES:
   If Err Then
       MsgBox Err.Number & "-" & Err.Description, vbInformation, "ERROR"
       Err = 0
       'cg.RollbackTrans
       Resume Next
    End If
End Sub

Private Sub Form_Load()
 MostrarForm Me, "C2"
 Mostrar_Data
 cAcepta.Enabled = False
 SSTab1.TabEnabled(1) = False
End Sub

Public Function Mostrar_Data()
  Dim sql As String
  Dim RS As New ADODB.Recordset
  Dim i As Integer
    
      sql = "SELECT a.vendedorcodigo as 'C�d.Vendedor'," & _
      "b.vendedornombres as Vendedor," & _
      "a.zonacodigo as 'C�d.Zona'," & _
      "c.zonadescripcion as Zona" & _
      " " & _
      "FROM  vt_zonavendedor a " & _
      "      JOIN  vt_vendedor b ON a.vendedorcodigo = b.vendedorcodigo " & _
      "      JOIN  vt_zona c ON a.zonacodigo = c.zonacodigo " & _
      "ORDER BY a.vendedorcodigo"
      
      Set RS = cn.Execute(sql)
      Set TDBGridProducto.DataSource = RS
    
      ' COMBO VENDEDOR:
      sql = "SELECT vendedorcodigo,vendedornombres " & _
      "FROM vt_vendedor "
      
      Set RS = cn.Execute(sql)
      If RS.RecordCount > 0 Then
        ReDim ArregloVendedor(0 To 1, 0 To RS.RecordCount - 1)
        Call fncLlenarArreglo_Combo(RS, cmbVendedor, ArregloVendedor, 1)
      End If
      ' COMBO ZONA:
      sql = "SELECT zonacodigo,zonadescripcion " & _
      "FROM vt_zona "
      
      Set RS = cn.Execute(sql)
      If RS.RecordCount > 0 Then
        ReDim ArregloZona(0 To 1, 0 To RS.RecordCount - 1)
        Call fncLlenarArreglo_Combo(RS, cmbZona, ArregloZona, 1)
      End If
    
  '    oCrystalReport.ReportFileName = RutaRep & "MantVendedorZona.rpt"
    
 TDBGridProducto.Refresh
 Set RS = Nothing
 SSTab1.Tab = 0
  
End Function


Private Function Validar_DatosNulos() As Boolean

                If cmbZona.ListIndex <> -1 _
                  And cmbVendedor.ListIndex <> -1 Then
                    Validar_DatosNulos = True
                    Exit Function
                End If

End Function


Private Sub SSTab1_Click(PreviousTab As Integer)
    SSTab1.TabEnabled(PreviousTab) = False
    cAcepta.Enabled = False
End Sub

Private Function Validar_CodigosDuplicados(Operacion As String, Optional filaorigen As Integer) As Boolean
Dim i As Integer
               
Validar_CodigosDuplicados = False
                    
 TDBGridProducto.MoveFirst
   Do Until TDBGridProducto.EOF
      If Operacion = "INSERT" Then
         If Trim(ArregloVendedor(0, cmbVendedor.ListIndex)) = _
                    Trim(TDBGridProducto.Columns(0).Text) Then
                        If Trim(ArregloZona(0, cmbZona.ListIndex)) = _
                           Trim(TDBGridProducto.Columns(2).Text) Then
                             Validar_CodigosDuplicados = True
                             Exit Function
                        End If
         End If
         
      ElseIf Operacion = "UPDATE" Then
         If Trim(ArregloVendedor(0, cmbVendedor.ListIndex)) = _
                 Trim(TDBGridProducto.Columns(0).Text) Then
                    If Trim(ArregloZona(0, cmbZona.ListIndex)) = _
                       Trim(TDBGridProducto.Columns(2).Text) _
                    And TDBGridProducto.Row <> i_filaorigen Then
                           Validar_CodigosDuplicados = True
                           Exit Function
                    End If
         End If
      End If
      TDBGridProducto.MoveNext
  Loop
    
End Function

Private Function MostrarOcultar_Botones(Valor As Boolean)
    frmbotones.Visible = Valor
End Function

Private Function fncSeleccionaCombo(ValorCodigo As String, Cbo As ComboBox, Arreglo As Variant)
Dim i As Integer
    For i = 0 To UBound(Arreglo, 2)
       If ValorCodigo = Arreglo(0, i) Then
         Cbo.ListIndex = i
         Exit Function
       End If
    Next i
End Function

Private Function fncLlenarArreglo_Combo(RS As Recordset, Cbo As ComboBox, Arreglo As Variant, dimensiones As Integer)
Dim i As Integer
Dim J As Integer

    i = 0
    Cbo.Clear
    Do Until RS.EOF
        Cbo.AddItem (Trim(RS(1)))
        For J = 0 To dimensiones
            Arreglo(J, i) = Trim(RS(J))
        Next J
        RS.MoveNext
        i = i + 1
    Loop
End Function

