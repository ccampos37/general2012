VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmMantAnalitico 
   Caption         =   "Analítico"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmbotones 
      Height          =   555
      Left            =   443
      TabIndex        =   3
      Top             =   5940
      Width           =   5805
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   330
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   330
         Index           =   1
         Left            =   1185
         TabIndex        =   7
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   330
         Index           =   2
         Left            =   2310
         TabIndex        =   6
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Salir"
         Height          =   330
         Index           =   4
         Left            =   4560
         TabIndex        =   5
         Top             =   165
         Width           =   1080
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Imprimir"
         Height          =   330
         Index           =   3
         Left            =   3435
         TabIndex        =   4
         Top             =   165
         Width           =   1080
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5910
      Left            =   23
      TabIndex        =   9
      Top             =   15
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   10425
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Consulta"
      TabPicture(0)   =   "frmMantAnalitico.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "lblNumReg"
      Tab(0).Control(2)=   "TDBGrid1"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmMantAnalitico.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "cCancela"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cAcepta"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame1 
         Height          =   4755
         Left            =   45
         TabIndex        =   12
         Top             =   375
         Width           =   6540
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
            Height          =   285
            Left            =   2700
            TabIndex        =   1
            Top             =   465
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   503
            XcodMaxLongitud =   6
            xcodwith        =   500
            NomTabla        =   "ct_entidad"
            ListaCampos     =   "entidadcodigo(1),entidadrazonsocial(1)"
            XcodCampo       =   "entidadcodigo"
            XListCampo      =   "entidadrazonsocial"
            ListaCamposDescrip=   "Codigo,Razon Social"
            ListaCamposText =   "entidadcodigo,entidadrazonsocial"
         End
         Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
            Height          =   285
            Left            =   2700
            TabIndex        =   0
            Top             =   180
            Width           =   3600
            _ExtentX        =   6350
            _ExtentY        =   503
            XcodMaxLongitud =   3
            xcodwith        =   500
            NomTabla        =   "ct_tipoanalitico"
            ListaCampos     =   "tipoanaliticocodigo(1),tipoanaliticodescripcion(1)"
            XcodCampo       =   "tipoanaliticocodigo"
            XListCampo      =   "tipoanaliticodescripcion"
            ListaCamposDescrip=   "Codigo,Descripcion"
            ListaCamposText =   "tipoanaliticocodigo,tipoanaliticodescripcion"
         End
         Begin VB.TextBox txtCod 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   2700
            TabIndex        =   2
            Top             =   780
            Width           =   3600
         End
         Begin VB.Label lbl 
            Caption         =   "Codigo Analitico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   165
            TabIndex        =   15
            Top             =   855
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Entidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   165
            TabIndex        =   14
            Top             =   555
            Width           =   2310
         End
         Begin VB.Label lbl 
            Caption         =   "Tipo Analitico"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   135
            TabIndex        =   13
            Top             =   225
            Width           =   2310
         End
      End
      Begin VB.CommandButton cAcepta 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   2055
         TabIndex        =   11
         Top             =   5325
         Width           =   1140
      End
      Begin VB.CommandButton cCancela 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   3495
         TabIndex        =   10
         Top             =   5325
         Width           =   1140
      End
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   5175
         Left            =   -74970
         TabIndex        =   16
         Top             =   360
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   9128
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
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
      Begin VB.Label lblNumReg 
         BackColor       =   &H00800000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lblNumReg"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   -69360
         TabIndex        =   18
         Top             =   5565
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Nº Registros"
         Height          =   270
         Left            =   -70320
         TabIndex        =   17
         Top             =   5580
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmMantAnalitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0


Dim modoinsert As Boolean
Dim modoedit As Boolean
Dim i_filaorigen As Integer
Dim rs As New ADODB.Recordset

Public Function MuestraDatos()
  Dim SQL As String
  
  SQL = "SELECT ct_analitico.analiticocodigo as Codigo, ct_analitico.tipoanaliticocodigo, ct_analitico.entidadcodigo, ct_entidad.entidadrazonsocial as Razon_Social,"
  SQL = SQL & "ct_tipoanalitico.tipoanaliticodescripcion as Tipo_Analitico "
  SQL = SQL & "FROM ct_analitico INNER JOIN ct_entidad ON ct_analitico.entidadcodigo = ct_entidad.entidadcodigo INNER JOIN   ct_tipoanalitico ON "
  SQL = SQL & "ct_analitico.tipoanaliticocodigo = ct_tipoanalitico.tipoanaliticocodigo "
  Set rs = VGcnx.Execute(SQL)
  Set TDBGrid1.DataSource = rs
  Call ConfiguraTdbgrid
  lblNumReg.Caption = rs.RecordCount
  SSTab1.Tab = 0
  
End Function

Private Sub cCancela_Click()
  SSTab1.TabEnabled(0) = True
  SSTab1.Tab = 0
  SSTab1.SetFocus
  frmbotones.Visible = True
  modoinsert = False
  modoedit = False
  i_filaorigen = -1
End Sub

Private Sub cmdBotones_Click(Index As Integer)
  Dim j As Integer
  Dim spos As Integer
  Dim SQL As String
  
  On Error GoTo x
  SSTab1.TabEnabled(1) = True
  
  Select Case Index
     Case 0   'nuevo
        SSTab1.Tab = 1
        Call LimpiarValores
        frmbotones.Visible = False
        modoinsert = True
        
     Case 1   'modificar
        If TDBGrid1.Row < 0 Then
          Exit Sub
        End If
        Call EditarValores
        modoedit = True
        SSTab1.Tab = 1
        frmbotones.Visible = False
        i_filaorigen = TDBGrid1.Row
      
     Case 2   'eliminar
       If MsgBox("Desea eliminar el registro?", vbYesNo + vbDefaultButton2, "AVISO") = vbYes Then
          SQL = "DELETE FROM CT_ANALITICO WHERE analiticocodigo='" & Trim(TDBGrid1.Columns(0).Value) & "'"
          VGcnx.Execute (SQL)
          Call MuestraDatos
       End If
        
     Case 3   'imprimir
       MDIPrincipal.cryRpt.Destination = crptToWindow
       MDIPrincipal.cryRpt.WindowState = crptMaximized
       MDIPrincipal.cryRpt.ReportFileName = App.Path & "\Reportes\rptMantAnalitico.rpt"
       MDIPrincipal.cryRpt.Connect = vgCADENAREPORT
       MDIPrincipal.cryRpt.DiscardSavedData = True
       MDIPrincipal.cryRpt.Action = 1
     
     Case 4  ' salir
       Unload Me
  End Select
   
x:
   If Err Then
      Err = 0
      Resume Next
   End If
   
End Sub

Sub EditarValores()
   txtCod.Text = Trim(TDBGrid1.Columns(0).Text)
   Ctr_Ayuda1.xclave = Trim(TDBGrid1.Columns(1).Text): Ctr_Ayuda1.Ejecutar
   Ctr_Ayuda2.xclave = Trim(TDBGrid1.Columns(2).Text): Ctr_Ayuda2.Ejecutar
End Sub

Public Function LimpiarValores()
  Ctr_Ayuda1.xclave = Empty: Ctr_Ayuda1.Ejecutar
  Ctr_Ayuda2.xclave = Empty: Ctr_Ayuda2.Ejecutar
  txtCod.Text = Empty
End Function

Private Sub cAcepta_Click()
  Dim SQL As String
  On Error GoTo x
  
  SSTab1.TabEnabled(0) = True
  
  If modoinsert = True Then
    'Validar Códigos Repetidos
    'Llamar al Metodo Grabar Tipo1,arreglo_valores,arrglo campos,arreglo tipodato
    SQL = "INSERT INTO CT_ANALITICO (analiticocodigo,entidadcodigo,tipoanaliticocodigo,usuariocodigo,fechaact) "
    SQL = SQL & "VALUES ('" & txtCod.Text & "','" & Ctr_Ayuda2.xclave & "','" & Ctr_Ayuda1.xclave & "','"
    SQL = SQL & VGusuario & "','" & Format(Date, "dd/mm/yyyy") & "')"
    VGcnx.BeginTrans
    VGcnx.Execute (SQL)
    VGcnx.CommitTrans
                  
  ElseIf modoedit = True Then
    'Validar Códigos Repetidos
    'Llamar al Metodo Grabar Tipo2,arreglo_valores,arreglo campos,arreglo tipo dato
    SQL = "UPDATE CT_ANALITICO SET "
  End If
  
  Call MuestraDatos
  frmbotones.Visible = True
  modoinsert = False: modoedit = False
  i_filaorigen = -1
  Exit Sub

x:
  If Err.Number = -2147217873 Then
    MsgBox "Esta intentando duplicar el Código de Analítico", vbInformation, Caption
  Else
    MsgBox "Error inesperado: " & Err.Number & " " & Err.Description
  End If
  VGcnx.RollbackTrans
     
End Sub

Private Sub Ctr_Ayuda2_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  If ColecCampos.Count > 0 Then
    txtCod.Text = Trim(Ctr_Ayuda1.xclave) & Trim(Ctr_Ayuda2.xclave)
    cAcepta.Enabled = ValidaDataIngreso()
  End If
End Sub

Private Sub Form_Load()
  Call ConfiguraForm
  Call MuestraDatos
End Sub

Sub ConfiguraForm()
  SSTab1.Tab = 0
  SSTab1.TabEnabled(1) = False
  Ctr_Ayuda1.conexion VGcnx
  Ctr_Ayuda2.conexion VGcnxMarfice
  cAcepta.Enabled = False
  lblNumReg.Caption = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set rs = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   SSTab1.TabEnabled(PreviousTab) = False
   cAcepta.Enabled = False
End Sub

Private Sub TDBGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    If rs.Sort = Empty Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
     ElseIf Right(rs.Sort, 3) = "asc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " desc"
     ElseIf Right(rs.Sort, 4) = "desc" Then
        rs.Sort = TDBGrid1.Columns.Item(ColIndex).DataField & " asc"
    End If
    Call ConfiguraTdbgrid
    TDBGrid1.Refresh
End Sub

Private Sub ConfiguraTdbgrid()
  Dim I As Integer
  Dim i_total As Integer
  Dim i_width As Integer
  TDBGrid1.Columns(1).Visible = False
  TDBGrid1.Columns(2).Visible = False
  TDBGrid1.Columns(0).Width = 10 * 100
  TDBGrid1.Columns(3).Width = 40 * 70
  TDBGrid1.Columns(4).Width = 40 * 70

End Sub

Function ValidaDataIngreso() As Boolean
  If Ctr_Ayuda1.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
   
  If Ctr_Ayuda2.xclave = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If
  
  If txtCod.Text = Empty Then
    ValidaDataIngreso = False
    Exit Function
  End If

  ValidaDataIngreso = True
End Function

Private Sub txtCod_GotFocus()
  txtCod.Text = Trim(Ctr_Ayuda1.xclave) & Trim(Ctr_Ayuda2.xclave)
End Sub
