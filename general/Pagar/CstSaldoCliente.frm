VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CstSaldoCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Saldo de Proveedor"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6375
      Index           =   1
      Left            =   210
      TabIndex        =   4
      Top             =   90
      Width           =   7095
      Begin VB.Frame Frame4 
         Height          =   930
         Left            =   2700
         TabIndex        =   8
         Top             =   5280
         Width           =   1980
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Cancelar"
            Height          =   690
            Index           =   12
            Left            =   1050
            Picture         =   "CstSaldoCliente.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   180
            Width           =   855
         End
         Begin VB.CommandButton cmdBotones 
            Caption         =   "&Acepta"
            Height          =   690
            Index           =   11
            Left            =   90
            Picture         =   "CstSaldoCliente.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   180
            Width           =   870
         End
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Left            =   1215
         TabIndex        =   0
         Top             =   315
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   582
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Ayuda de Clientes"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
         Requerido       =   0   'False
      End
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   4275
         Left            =   90
         TabIndex        =   6
         Top             =   960
         Width           =   6885
         Begin MSFlexGridLib.MSFlexGrid MGrid1 
            Height          =   4305
            Left            =   30
            TabIndex        =   7
            Top             =   0
            Width           =   6855
            _ExtentX        =   12091
            _ExtentY        =   7594
            _Version        =   393216
            FixedCols       =   0
            GridLines       =   0
            BorderStyle     =   0
            Appearance      =   0
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   0
         Left            =   150
         TabIndex        =   5
         Top             =   375
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar Panel 
      Align           =   2  'Align Bottom
      Height          =   348
      Left            =   0
      TabIndex        =   3
      Top             =   6864
      Width           =   7572
      _ExtentX        =   13361
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "CstSaldoCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBotones_Click(Index As Integer)
  Dim rb As New ADODB.Recordset
  Dim J As Integer
  Dim wsumasol, wsumadol As Double
  Select Case Index
    Case 11 'consultar datos
      Call Cargar_Flex
      
      Set rb = VGCNx.Execute("select documentocargo,tdocumentodescripcion,tdocumentotipo," & _
                        " round(sum( case monedacodigo when '02' then  isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) else 0 end),2),round(sum( case monedacodigo when '01' then  isnull(cargoapeimpape,0)-isnull(cargoapeimppag,0) else 0 end),2)" & _
                        " From cp_cargo " & _
                        " Inner join cp_tipodocumento " & _
                        " on cp_cargo.documentocargo=cp_tipodocumento.tdocumentocodigo " & _
                        " where clientecodigo='" & Ctr_Ayuda1.xclave & "' and (cargoapeflgreg<>'1' or cargoapeflgreg is null) " & _
                        " group by documentocargo,tdocumentodescripcion,tdocumentotipo")
    
      If rb.RecordCount > 0 Then
         rb.MoveFirst
         wsumasol = 0: wsumadol = 0
         Do Until rb.EOF
            MGrid1.Rows = MGrid1.Rows + 1
            MGrid1.Row = MGrid1.Rows - 1
            MGrid1.Col = 0: MGrid1.CellAlignment = 0: MGrid1.TextMatrix(MGrid1.RowSel, 0) = Trim$(rb.Fields(0))
            MGrid1.Col = 1: MGrid1.CellAlignment = 0: MGrid1.TextMatrix(MGrid1.RowSel, 1) = Trim$(rb.Fields(1))
            MGrid1.Col = 2: MGrid1.CellAlignment = 1: MGrid1.CellFontSize = 10: MGrid1.TextMatrix(MGrid1.RowSel, 2) = IIf(rb.Fields(2) = "A", "(+)", "(-)")
            If rb.Fields("tdocumentotipo") = "A" Then
               MGrid1.Col = 3: MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 3) = "(" & Numero(rb.Fields(3)) & ")"
               MGrid1.Col = 4: MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 4) = "(" & Numero(rb.Fields(4)) & ")"
               wsumasol = wsumasol - rb.Fields(4)
               wsumadol = wsumadol - rb.Fields(3)
            ElseIf rb.Fields("tdocumentotipo") = "C" Then
               MGrid1.Col = 3: MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 3) = Numero(rb.Fields(3))
               MGrid1.Col = 4: MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 4) = Numero(rb.Fields(4))
               wsumasol = wsumasol + rb.Fields(4)
               wsumadol = wsumadol + rb.Fields(3)
            End If
            rb.MoveNext
         Loop
         MGrid1.Rows = MGrid1.Rows + 2
         MGrid1.Row = MGrid1.Rows - 1
         MGrid1.Col = 0: MGrid1.CellBackColor = RGB(0, 100, 155): MGrid1.CellForeColor = RGB(0, 255, 0): MGrid1.TextMatrix(MGrid1.RowSel, 0) = "TOTALES"
         MGrid1.Col = 1: MGrid1.CellBackColor = RGB(0, 100, 155): MGrid1.CellForeColor = RGB(0, 255, 0): MGrid1.TextMatrix(MGrid1.RowSel, 1) = "GENERALES "
         MGrid1.Col = 2: MGrid1.CellBackColor = RGB(0, 100, 155): MGrid1.CellForeColor = RGB(0, 255, 0): MGrid1.TextMatrix(MGrid1.RowSel, 2) = ""
         MGrid1.Col = 3: MGrid1.CellBackColor = RGB(0, 100, 155): MGrid1.CellForeColor = RGB(0, 255, 0): MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 3) = Numero(wsumadol)
         MGrid1.Col = 4: MGrid1.CellBackColor = RGB(0, 100, 155): MGrid1.CellForeColor = RGB(0, 255, 0): MGrid1.CellAlignment = 8: MGrid1.TextMatrix(MGrid1.RowSel, 4) = Numero(wsumasol)
         
      End If
      rb.Close
      
    Case 12 ' salir
      Unload Me
        
  End Select
  Set rb = Nothing
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
  If Len(Trim$(Ctr_Ayuda1.xclave)) > 0 Then
    Call cmdBotones_Click(11)
  End If
End Sub

Private Sub Form_Load()
  MostrarForm Me, "C"
  Call Ctr_Ayuda1.conexion(VGCNx)
  
  Call Cargar_Flex
End Sub

Public Sub Cargar_Flex()
    MGrid1.Clear
    MGrid1.Rows = 1: MGrid1.Cols = 5
    MGrid1.Row = 0: MGrid1.ColWidth(0) = 1000: MGrid1.TextMatrix(0, 0) = "T/Doc."
    MGrid1.Row = 0: MGrid1.ColWidth(1) = 2400: MGrid1.TextMatrix(0, 1) = "Documentos"
    MGrid1.Row = 0: MGrid1.ColWidth(2) = 400: MGrid1.TextMatrix(0, 2) = "C/A"
    MGrid1.Row = 0: MGrid1.ColWidth(3) = 1510: MGrid1.TextMatrix(0, 3) = "Total/Dolares"
    MGrid1.Row = 0: MGrid1.ColWidth(4) = 1510: MGrid1.TextMatrix(0, 4) = "Total/Soles"
End Sub
