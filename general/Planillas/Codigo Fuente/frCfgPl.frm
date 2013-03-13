VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frCfgPl 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración de Vista de Planillas"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4635
   Icon            =   "frCfgPl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2955
      Top             =   3525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCfgPl.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   105
      TabIndex        =   1
      Top             =   3660
      Width           =   1290
   End
   Begin MSComctlLib.ListView LColumnas 
      Height          =   3420
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6033
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4095
      Picture         =   "frCfgPl.frx":0796
      Top             =   3600
      Width           =   480
   End
End
Attribute VB_Name = "frCfgPl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim RSCOLPL As New ADODB.Recordset
    Dim ANOMPL()
    Dim ACODPL(), X As Integer
    ANOMPL = Array("MES DE PROCESO", "TIPO DE PLANILLA", "NUMERO DE BOLETA DE REMUNERACIONES", "CÓDIGO DEL TRABAJADOR", "NOMBRES DEL TRABAJADOR", "TIPO DE TRABAJADOR", "FECHA DE INGRESO", "SITUACIÓN DEL TRABAJADOR", "CÓDIGO DEL CENTRO DE COSTO", "NOMBRE DEL CENTRO DE COSTO", "DEPARTAMENTO DEL TRABAJADOR", "CARGO DEL TRABAJADOR", "REMUNERACIÓN BÁSICA", "FONDO DE PENSIONES", "FECHA DE CESE DEL TRABAJADOR", "CODIGO DE SCTR", "NÚMERO DE RUC DE E.P.S.")
    ACODPL = Array("MES", "TIPOPLANILLA", "INUMBOL", "CODTRAB", "NOMBRES", "TIPOTRAB", "FECHAING", "SITUACION", "CCOSTO", "CENTROCOSTO", "DEPARTAMENTO", "CARGO", "BASICO", "FONDOPENS", "FECHACESE", "CODSCTR", "EPS")
    If Not ExisteTabla("VERCOLPL") Then
        DBSYSTEM.Execute "CREATE TABLE VERCOLPL (CODIGO VARCHAR(15), VALOR BIT)"
    End If
    DBSYSTEM.Execute "DELETE FROM VERCOLPL"
    RSCOLPL.Open REGSISTEMA.TABLAPLAN, DBSYSTEM, adOpenKeyset
    For X = 1 To RSCOLPL.Fields.Count
        DBSYSTEM.Execute "INSERT INTO VERCOLPL (CODIGO, VALOR) VALUES ('" & Trim$(RSCOLPL.Fields(X - 1).Name) & "'," & IIf(frVerPlan.DataPlan.Columns(RSCOLPL.Fields(X - 1).Name).Visible, 1, 0) & ")"
    Next
    RSCOLPL.Close
    Dim XITEM As ListItem, XNOMB As String
    For X = 0 To 16
        Set XITEM = LColumnas.ListItems.Add(, ACODPL(X), ACODPL(X), , 1)
        XITEM.SubItems(1) = ANOMPL(X)
    Next
    RSCOLPL.Open "SELECT CODIGO, NOMBRE FROM COLUMPL ORDER BY INDICE", DBSYSTEM, adOpenStatic
    Do While Not RSCOLPL.EOF
        Set XITEM = LColumnas.ListItems.Add(, RSCOLPL!CODIGO, RSCOLPL!CODIGO, , 1)
        XITEM.SubItems(1) = RSCOLPL!NOMBRE
        RSCOLPL.MoveNext
    Loop
    RSCOLPL.Close
    RSCOLPL.Open "VERCOLPL", DBSYSTEM, adOpenKeyset
    Do While Not RSCOLPL.EOF
        LColumnas.ListItems(Trim$(RSCOLPL!CODIGO)).Checked = IIf(RSCOLPL!VALOR = 1, True, False)
        RSCOLPL.MoveNext
    Loop
    Set RSCOLPL = Nothing
End Sub

Private Sub LCOLUMNAS_ITEMCHECK(ByVal ITEM As MSComctlLib.ListItem)
    frVerPlan.DataPlan.Columns(ITEM.KEY).Visible = ITEM.Checked
    DBSYSTEM.Execute "UPDATE VERCOLPL SET VALOR=" & IIf(ITEM.Checked, 1, 0) & " WHERE CODIGO='" & ITEM.KEY & "'"
End Sub

